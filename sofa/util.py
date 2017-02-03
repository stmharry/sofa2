# -*- coding: utf-8 -*-

import bs4
import collections
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

DEBUG = True


class Document(object):
    def __init__(self,
                 ids,
                 checked,
                 source,
                 source_no,
                 receivers,
                 receive_datetime,
                 subject,
                 num_attachments):

        self.ids = ids
        self.checked = checked and False # DEBUG
        self.source = source
        self.source_no = source_no
        self.receivers = receivers
        self.receive_datetime = receive_datetime
        self.subject = subject
        self.num_attachments = num_attachments


class Manager(object):
    SUCCESS = 0

    @staticmethod
    def time_str(time):
        return '{:d}-{:02d}-{:02d}'.format(
            time.year - 1911,
            time.month,
            time.day,
        )

    @staticmethod
    def code_str(item):
        return u'{item.code_no} {item.code_nm}'.format(item=item)

    @staticmethod
    def document_str(document, status):
        if status == Manager.SUCCESS:
            str_ = u'公文 {:s} 已掛號為 {:d} 號，承辦人為 {:s}'.format(
                document.source_no,
                document.receive_no,
                document.user_nm,
            )
        return str_

    def __init__(self,
                 eclient,
                 connection):

        self.eclient = eclient
        self.connection = connection
        self.conductors = connection.select(
            from_='conductor',
            fields={
                'user_nm': Connection.rtrim('user_nm'),
            },
        )
        self.conductors.loc[self.conductors.user_nm == u'莊俊傑', 'path'] = u'\\\\隊本部收發\\收發\\公文附件\\俊傑'  # DEBUG
        self.conductors.loc[self.conductors.user_nm == u'粘銘進', 'path'] = u'\\\\隊本部收發\\收發\\公文附件\\銘進'  # DEBUG
        self.conductors.loc[self.conductors.user_nm == u'林舜欽', 'path'] = u'\\\\隊本部收發\\收發\\公文附件\\舜欽'  # DEBUG
        self.conductors.loc[self.conductors.user_nm == u'林鴻慶', 'path'] = u'\\\\隊本部收發\\收發\\公文附件\\鴻慶'  # DEBUG
        self.conductors.loc[self.conductors.user_nm == u'曾明欽', 'path'] = u'\\\\隊本部收發\\收發\\公文附件\\曾明欽'  # DEBUG
        self.conductors.loc[self.conductors.user_nm == u'劉晃', 'path'] = u'\\\\隊本部收發\\收發\\公文附件\\劉晃'  # DEBUG
        self.conductors.loc[self.conductors.user_nm == u'蔣招祺', 'path'] = u'\\\\隊本部收發\\收發\\公文附件\\招祺'  # DEBUG

        self.print_path = u'\\\\隊本部收發\\收發\\公文附件\\列印'  # DEBUG

        self.secrets = connection.select(
            from_='secret',
            fields={
                'code_no': Connection.rtrim('code_no'),
                'code_nm': Connection.rtrim('code_nm'),
            },
        )
        self.papers = connection.select(
            from_='paper',
            fields={
                'code_no': Connection.rtrim('code_no'),
                'code_nm': Connection.rtrim('code_nm'),
            },
        )
        self.books = connection.select(
            from_='book',
            fields={
                'code_no': Connection.rtrim('code_no'),
                'code_nm': Connection.rtrim('code_nm'),
            },
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

    def receive(self,
                source_word='',
                source_number='',
                source='',
                subject='',
                start_datetime=None,
                end_datetime=None):

        now = datetime.datetime.now()
        start_datetime = start_datetime or now
        end_datetime = end_datetime or (now - datetime.timedelta(weeks=1))

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
            'sHour': start_datetime.hour,
            'endDate': Manager.time_str(end_datetime),
            'eHour': end_datetime.hour,
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

        documents_by_source_no = collections.OrderedDict()

        for (tr0, tr1) in tr_batches:
            tds = tr0.find_all('td')
        
            if u'收文完成' not in tds[3].contents[1].string:
                continue

            id_ = int(urlparse.parse_qs(urlparse.urlparse(tr0['linkto']).query)['dilistid'][0])
            source_no = u'{:s}字第{:d}號'.format(tds[5].string, int(tds[6].contents[1].string))
            receiver = tds[9].contents[2].string

            if source_no in documents_by_source_no:
                document = documents_by_source_no[source_no]

                document.ids.append(id_)
                document.receivers.append(receiver)
            else:
                document = Document(
                    ids=[id_],
                    checked=tds[1].input.has_attr('checked'),
                    source=tds[4].contents[2].string,
                    source_no=source_no,
                    receivers=[receiver],
                    receive_datetime=datetime.datetime.strptime(tds[8].string.split(' ')[0], '%Y/%m/%d'),
                    subject=tr1.contents[1].contents[2].string.strip().split(u'：', 1)[1],
                    num_attachments=int(tds[7].string),
                )
                documents_by_source_no[source_no] = document

        return documents_by_source_no

    def receive_detail(self, document):
        document.attachments = {}
        for (num, id_) in enumerate(document.ids):
            params = {
                'menuCode': 'RECVQRY',
                'detail': 'showdetail',
                'dilistid': id_,
                'listid': '',
            }

            r = self.eclient.get('webeClient/main.php', params=params)
            soup = bs4.BeautifulSoup(r.content, 'html.parser')
            table = soup.find('table', id='Table1')
            trs = table.find_all('tr')

            if num == 0:
                tds = trs[1].find_all('td')
                document.paper_nm = tds[3].string
                document.speed_nm = tds[1].string

                as_ = trs[4].find_all('a')
                for a_ in as_[1:]:
                    if not a_.string.endswith('.di') and not a_.string.endswith('.sw'):
                        document.attachments[a_.string] = a_['href']

            input_ = soup.find('input', value=u'下載PDF')
            match = re.search('(\'(?P<url>..*)\')', input_['onclick'])
            pdf_name = u'{:s}.pdf'.format(document.subject[:8])
            document.attachments[pdf_name] = match.group('url')


    def set_checked(self, document):
        document.checked = not document.checked

        if DEBUG:
            return

        params = {
            '_': int(time.time() * 1000),
            'menuCode': 'RECVQRY',
            'showhtml': 'empty',
            'action': 'settag',
            'dilistid': document.id_,
            'tagvalue': int(document.checked),
        }

        r = self.eclient.get('webeClient/main.php', params=params)

    def save_files(self, document):
        document.receive_no = self.receive_no

        if DEBUG:
            return

        conductor = self.conductors[self.conductors.user_nm == document.user_nm].iloc[0]
        attachment_dir = os.path.join(conductor.path, '{:04d}'.format(document.receive_no))

        if not os.path.isdir(attachment_dir):
            os.mkdir(attachment_dir)
        
        now = time.time()
        for (num_attachment, (name, url)) in enumerate(document.attachments.items()):
            r = self.eclient.get('webeClient/{:s}'.format(url), stream=True)

            attachment_path = os.path.join(
                attachment_dir, 
                '{:d}_{:d}_{:s}'.format(document.receive_no, num_attachment, name),
            )
            with open(attachment_path, 'wb') as f:
                shutil.copyfileobj(r.raw, f)

            print_path = os.path.join(self.print_path, '{:.0f}_{:s}'.format(now, name))
            shutil.copyfile(attachment_path, print_path)

    def to_archive(self, document):
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
            self.sno = 1
            self.receive_no = 1
        else:
            self.sno = archives.sno.max() + 1
            self.receive_no = archives.receive_no.max() + 1

        archive = self.archive_default.copy().assign(
            ymm_year=now.year,
            sno=self.sno,
            receive_no=self.receive_no,
            receive_date='{:%Y/%m/%d}'.format(now),
            source=document.source,
            source_no=document.source_no,
            paper=Manager.code_str(
                item=self.papers[self.papers.code_nm == document.paper_nm].iloc[0],
            ),
            user_nm=document.user_nm,
            subject=document.subject,
            ymm_month=now.month,
        )

        return archive

class eClient(requests.Session):
    def __init__(self,
                 server,
                 userid,
                 passwd):

        super(eClient, self).__init__()
        self.server = server

        self.headers.update({'referer': ''})
        self.post(
            'webeClient/menu.php',
            data={
                'login': 1,
                'userid': userid,
                'passwd': passwd,
            },
        )

    def route(self, url):
        url = urlparse.urljoin(self.server, url)
        return url

    def get(self, url, *args, **kwargs):
        return super(eClient, self).get(self.route(url), *args, **kwargs)

    def post(self, url, *args, **kwargs):
        return super(eClient, self).post(self.route(url), *args, **kwargs)


class Connection(pypyodbc.Connection):
    @staticmethod
    def rtrim(field):
        return 'rtrim(' + field + ')'

    @staticmethod
    def sentence(strs, func=unicode, sep=',', begin='', end='', default=''):
        begin = begin + ' '
        sep = ' ' + sep + ' '
        end = ' ' + end
        return (begin + sep.join(map(func, strs)) + end) if strs else default

    def select(self,
               from_,
               top=None,
               fields=['*'],
               wheres=[],
               order_bys=[]):

        if isinstance(fields, dict):
            field_names = fields.keys()
            field_values = fields.values()
        else:
            field_names = None
            field_values = fields

        query = (
            u'select {top}{field} '
            u'from {from_} '
            u'{where} '
            u'{order_by} '
        ).format(
            top='' if top is None else 'top {} '.format(top),
            field=Connection.sentence(field_values),
            from_=from_,
            where=Connection.sentence(wheres, sep='and', begin='where'),
            order_by=Connection.sentence(order_bys, begin='order by'),
        )

        print(query)

        df = pd.read_sql(query, con=self)
        if field_names is not None:
            df.columns = field_names
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
                    func=lambda value: '\'' + unicode(value) + '\'',
                    begin='(',
                    end=')',
                ) for row in df.itertuples(index=False)
            ]),
        )

        print(query)

        if DEBUG:
            return

        cursor = self.cursor()
        cursor.execute(query)
        cursor.close()
        self.commit()