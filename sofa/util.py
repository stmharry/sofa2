# -*- coding: utf-8 -*-

import bs4
import collections
import configobj
import cStringIO
import datetime
import itertools
import pandas as pd
import pypyodbc
import os
import re
import requests
import shutil
import time
import urlparse

pd.set_option('display.unicode.east_asian_width', True)

FLAG_CHECKED = False
FLAG_INSERT = False
FLAG_SAVE = False


class Manager(object):
    PRINT_ONLY = u'列印'

    @staticmethod
    def time_str(t):
        return '{:d}-{:02d}-{:02d}'.format(
            t.year - 1911,
            t.month,
            t.day,
        )

    @staticmethod
    def code_str(item):
        return u'{item.code_no} {item.name}'.format(item=item)

    def __init__(self,
                 eclient,
                 connection,
                 config_path):

        self.eclient = eclient
        self.connection = connection
        self.config = configobj.ConfigObj(config_path, encoding='utf-8')
        self.debug_messages = []
        self.alerts = []

        self.print_dir = self.config['print_dir']
        if not os.path.isdir(self.print_dir):
            os.mkdir(self.print_dir)
        self.conductors = connection.select(
            from_='conductor',
            fields={
                'user_nm': Connection.rtrim('user_nm'),
            },
            index_col='user_nm',
        )
        for (user_nm, attachment_dir) in self.config['attachment_dir'].items():
            self.conductors.loc[user_nm, 'attachment_dir'] = attachment_dir
        
        self.secrets = connection.select(
            from_='secret',
            fields={
                'code_no': Connection.rtrim('code_no'),
                'code_nm': Connection.rtrim('code_nm'),
            },
            index_col='code_nm',
        )
        self.papers = connection.select(
            from_='paper',
            fields={
                'code_no': Connection.rtrim('code_no'),
                'code_nm': Connection.rtrim('code_nm'),
            },
            index_col='code_nm',
        )
        self.books = connection.select(
            from_='book',
            fields={
                'code_no': Connection.rtrim('code_no'),
                'code_nm': Connection.rtrim('code_nm'),
            },
            index_col='code_nm',
        )

        secret_default = self.secrets[self.secrets.code_no == '1'].iloc[0]
        paper_default = self.papers[self.papers.code_no == '21'].iloc[0]
        book_default = self.books[self.books.code_no == '1'].iloc[0]

        self.archive_default = pd.DataFrame([{
            'secret': Manager.code_str(secret_default),
            'paper': Manager.code_str(paper_default),
            'book': Manager.code_str(book_default),
            'speed': u'普通件',
            'archive_no': 1,
            'process_days': 0.5,
            'measure': u'頁',
            'ymm_user': u'系統管理員',
        }])

    def process(self, document, conductor):
        document.user_nm = conductor
        document.print_only = (conductor == Manager.PRINT_ONLY)

        self.receive_detail(document)
        if not document.print_only:
            self.insert(document)
        self.save_as_print(document)
        if not document.print_only:
            self.save_as_attachment(document)
        self.success(document)

    def receive(self,
                source_word='',
                source_number='',
                source='',
                subject='',
                start_datetime=None,
                end_datetime=None):

        now = datetime.datetime.now()
        start_datetime = start_datetime or (now - datetime.timedelta(weeks=1))
        end_datetime = end_datetime or now

        params = {
            'menuCode': 'RECVQRY',
            'detail': 'show',
            'recvname': '',
            'doc_word': source_word,
            'doc_no': source_number,
            'doc_title': subject,
            'full_compare': 0,
            'senderid': '',
            'sendername': source,
            'query_time': 'create_time',
            'startDate': Manager.time_str(start_datetime),
            'sHour': '00',
            'endDate': Manager.time_str(end_datetime),
            'eHour': '23',
            'noteonly': 2,
            'pn': 0,
            'ifquery_recvqry': 1,
            'full_compare_send': '',
        }
        r = self.eclient.get('webeClient/main.php', params=params)

        soup = bs4.BeautifulSoup(r.content, 'html.parser')
        table = soup.find('table', id='Table2')
        trs = table.find_all(
            'tr',
            class_='openMenu',
        )
        tr_batches = itertools.izip_longest(*[iter(trs)] * 2)

        document_by_source_no = collections.OrderedDict()

        for (tr0, tr1) in tr_batches:
            tds = tr0.find_all('td')

            if u'收文完成' not in tds[3].contents[1].string:
                continue

            source_no = u'{:s}字第{:d}號'.format(tds[5].string, int(tds[6].contents[1].string))
            receive_datetime_str = tds[8].string.replace(u'上午', 'AM').replace(u'下午', 'PM')

            if source_no not in document_by_source_no:
                document_by_source_no[source_no] = Document(
                    source=tds[4].contents[2].string,
                    source_no=source_no,
                    receive_datetime=datetime.datetime.strptime(receive_datetime_str, '%Y/%m/%d %p %I:%M:%S'),
                    subject=tr1.contents[1].contents[2].string.strip().split(u'：', 1)[1],
                    num_attachments=int(tds[7].string),
                )

            document_by_source_no[source_no].add_branch(
                id_=int(urlparse.parse_qs(urlparse.urlparse(tr0['linkto']).query)['dilistid'][0]),
                checked=tds[1].input.has_attr('checked'),
                receiver=tds[9].contents[2].string,
            )

        return document_by_source_no

    def receive_detail(self, document):
        attachment_names = set()

        if len(document.branches) == 1:
            name_str = u'{document.source_no:s}.pdf'
        else:
            name_str = u'{document.source_no:s}_{branch.receiver:s}.pdf'

        for branch in document.branches:
            params = {
                'menuCode': 'RECVQRY',
                'detail': 'showdetail',
                'dilistid': branch.id_,
                'listid': '',
            }

            r = self.eclient.get('webeClient/main.php', params=params)
            soup = bs4.BeautifulSoup(r.content, 'html.parser')
            table = soup.find('table', id='Table1')
            trs = table.find_all('tr')

            input_ = soup.find('input', value=u'下載PDF')
            match = re.search('(\'(?P<url>..*)\')', input_['onclick'])

            document.add_attachment(
                name=name_str.format(document=document, branch=branch),
                buf=self.eclient.download('webeClient/{:s}'.format(match.group('url'))),
                is_main=branch.is_main and not document.print_only,
            )

            as_ = trs[4].find_all('a')
            for a_ in as_[1:]:
                name = a_.string
                if name.endswith('.di') or name.endswith('.sw') or name in attachment_names:
                    continue

                document.add_attachment(
                    name=name,
                    buf=self.eclient.download('webeClient/{:s}'.format(a_['href'])),
                    is_main=False,
                )

        tds = trs[1].find_all('td')
        document.paper_nm = tds[3].string
        document.speed_nm = tds[1].string
                
    def insert(self, document):
        now = datetime.datetime.now()
        archives = self.connection.select(
            from_='archive',
            fields={
                'sno': 'sno',
                'receive_no': 'convert(int, receive_no)',
            },
            wheres=['Year(receive_date) = {:d}'.format(now.year)],
            order_bys=['sno desc'],
        )
        if archives.empty:
            document.sno = 1
            document.receive_no = 1
        else:
            document.sno = archives.sno.max() + 1
            document.receive_no = archives.receive_no.max() + 1

        archive = self.archive_default.copy().assign(
            ymm_year=now.year,
            sno=document.sno,
            receive_no=document.receive_no,
            receive_date='{:%Y/%m/%d}'.format(now),
            source=document.source,
            source_no=document.source_no,
            paper=Manager.code_str(
                item=self.papers[self.papers.index == document.paper_nm].iloc[0],
            ),
            user_nm=document.user_nm,
            subject=document.subject,
            ymm_month=now.month,
        )

        if not FLAG_INSERT:
            return

        self.connection.insert(
            archive,
            into='archive',
        )

    def save_as_print(self, document):
        if document.receive_no is None:
            no_str = '{document.source_no:s}'
        else:
            no_str = '{document.receive_no:04d}'

        for (num_attachment, attachment) in enumerate(document.attachments):
            path = os.path.join(
                self.print_dir,
                (no_str + u'_附件{num_attachment:d}_{attachment.name:s}').format(
                    document=document,
                    num_attachment=num_attachment,
                    attachment=attachment,
                ),
            )

            if not FLAG_SAVE:
                continue
            
            attachment.save(path)

    def save_as_attachment(self, document):
        conductor = self.conductors.loc[document.user_nm]
        document.attachment_dir = os.path.join(conductor.attachment_dir, '{:04d}'.format(document.receive_no))

        for attachment in document.attachments:
            path = os.path.join(
                document.attachment_dir,
                attachment.name,
            )

            if not FLAG_SAVE:
                continue

            attachment.save(path)

    def success(self, document):
        document.checked = True
        for branch in document.branches:
            branch.checked = True

        if not FLAG_SAVE:
            return

        for branch in document.branches:
            params = {
                '_': int(time.time() * 1000),
                'menuCode': 'RECVQRY',
                'showhtml': 'empty',
                'action': 'settag',
                'dilistid': branch.id_,
                'tagvalue': int(branch.checked),
            }

            self.eclient.get('webeClient/main.php', params=params)


class Document(object):
    class _Branch(object):
        def __init__(self, id_, checked, receiver):
            self.id_ = id_
            self.checked = checked and FLAG_CHECKED  # DEBUG
            self.receiver = receiver
            self.is_main = re.search(u'(第三|本)大隊[^中]*$', receiver) is not None

    class _Attachment(object):
        def __init__(self, name, buf, is_main):
            self.name = name
            self.buf = buf
            self.is_main = is_main

        def save(self, path):
            dir_ = os.path.dirname(path)
            if not os.path.isdir(dir_):
                os.makedirs(dir_)

            with open(path, 'wb') as f:
                self.buf.seek(0)
                shutil.copyfileobj(self.buf, f)

    def __init__(self,
                 source,
                 source_no,
                 receive_datetime,
                 subject,
                 num_attachments):

        self.branches = []
        self.attachments = []

        self.checked = True and FLAG_CHECKED  # DEBUG
        self.sno = None
        self.source = source
        self.source_no = source_no
        self.source_is_self = re.search(u'保七三大[^中]*字', source_no) is not None
        self.receive_no = None
        self.receive_datetime = receive_datetime
        self.subject = subject
        self.num_attachments = num_attachments

    def add_branch(self, id_, checked, receiver):
        self.branches.append(
            self._Branch(
                id_=id_,
                checked=checked,
                receiver=receiver,
            )
        )
        self.checked = self.checked and checked

    def add_attachment(self, name, buf):
        self.attachments.append(
            self._Attachment(
                name=name,
                buf=buf,
            )
        )


class eClient(requests.Session):
    def __init__(self,
                 server,
                 userid,
                 passwd):

        super(eClient, self).__init__()
        self.server = server
        self.userid = userid
        self.passwd = passwd

        self.headers.update({'referer': ''})

    def route(self, url):
        return urlparse.urljoin(self.server, url)

    def request(self, method, url, **kwargs):
        return super(eClient, self).request(method, self.route(url), **kwargs)

    def get(self, url, params=None, **kwargs):
        return self.request('get', url, params=params, **kwargs)

    def post(self, url, data=None, json=None, **kwargs):
        return self.request('post', url, data=data, json=json, **kwargs)

    def login(self):
        self.post(
            'webeClient/menu.php',
            data={
                'login': 1,
                'userid': self.userid,
                'passwd': self.passwd,
            },
        )

    def download(self, url):
        r = self.get(
            url, 
            stream=True,
        )
        buf = cStringIO.StringIO()
        shutil.copyfileobj(r.raw, buf)
        return buf


class Connection(pypyodbc.Connection):
    @staticmethod
    def rtrim(field):
        return 'rtrim(' + field + ')'

    @staticmethod
    def pad(str_, func=unicode, pad=' '):
        return pad + func(str_) + pad

    @staticmethod
    def sentence(strs, func=unicode, sep=',', begin='', end='', default=''):
        return (
            begin +
            Connection.pad(
                Connection.pad(sep).join(map(func, strs))
            ) +
            end
        ) if strs else default

    def select(self,
               from_,
               top=None,
               fields=None,
               index_col=None,
               wheres=None,
               order_bys=None):

        fields = fields or ['*']
        wheres = wheres or []
        order_bys = order_bys or []

        if isinstance(fields, dict):
            field_names = fields.keys()
            field_values = fields.values()
        else:
            field_names = None
            field_values = fields

        query = (
            u'select {top} {field} ' +
            u'from {from_} ' +
            u'{where} ' +
            u'{order_by} '
        ).format(
            top='' if top is None else 'top {:s} '.format(top),
            field=Connection.sentence(field_values),
            from_=from_,
            where=Connection.sentence(wheres, sep='and', begin='where'),
            order_by=Connection.sentence(order_bys, begin='order by'),
        )

        df = pd.read_sql(query, con=self)

        if field_names is not None:
            df.columns = field_names
        if index_col is not None:
            df.set_index(index_col, inplace=True)
        return df

    def insert(self,
               df,
               into):

        query = (
            u'insert '
            u'into {into} {field} '
            u'values {value} '
        ).format(
            into=into,
            field=Connection.sentence(list(df.columns), begin='(', end=')'),
            value=Connection.sentence([
                Connection.sentence(
                    row,
                    func=lambda str_: Connection.pad(str_, pad='\''),
                    begin='(',
                    end=')',
                ) for row in df.itertuples(index=False)
            ]),
        )

        cursor = self.cursor()
        cursor.execute(query)
        cursor.close()
        self.commit()