# -*- coding: utf-8 -*-

import bs4
import datetime
import itertools
import pandas as pd
import pypyodbc
import requests
import urlparse

pd.set_option('display.unicode.east_asian_width', True)


class Document(object):
    def __init__(self,
                 id_,
                 checked,
                 source,
                 source_no,
                 receiver,
                 receive_datetime,
                 subject,
                 secret_nm=None,  # TODO
                 paper_nm=None,  # TODO
                 book_nm=None,  # TODO
                 speed_nm=None,
                 attachments=None,
                 user_nm=None):

        self.id_ = id_
        self.checked = checked
        self.source = source
        self.source_no = source_no
        self.receiver = receiver
        self.receive_datetime = receive_datetime
        self.subject = subject

        self.secret_nm = secret_nm
        self.paper_nm = paper_nm
        self.book_nm = book_nm
        self.speed_nm = speed_nm
        self.attachments = attachments
        self.user_nm = user_nm


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
        return urlparse.urljoin(self.server, url)

    def get(self, url, *args, **kwargs):
        return super(eClient, self).get(self.route(url), *args, **kwargs)

    def post(self, url, *args, **kwargs):
        return super(eClient, self).post(self.route(url), *args, **kwargs)

    def receive(self,
                source_word='',
                source_number='',
                source_id='',
                source='',
                subject='',
                full_compare=0,
                query_time='create_time',
                start_date=None,
                start_hour=0,
                end_date=None,
                end_hour=23):

        now = datetime.datetime.now()
        now_str = '{:d}-{:02d}-{:02d}'.format(
            now.year - 1911,
            now.month,
            now.day,
        )

        params = {
            'menuCode': 'RECVQRY',
            'detail': 'show',
            'recvname': '',
            'doc_word': source_word,
            'doc_no': source_number,
            'doc_title': subject,
            'full_compare': full_compare,
            'senderid': source_id,
            'sendername': source,
            'query_time': query_time,
            'startDate': start_date or now_str,
            'sHour': start_hour,
            'endDate': end_date or now_str,
            'eHour': end_hour,
            'noteonly': 2,
            'pn': 0,
            'ifquery_recvqry': 1,
            'full_compare_send': '',
        }
        r = self.get('webeClient/main.php', params=params)

        soup = bs4.BeautifulSoup(r.content, 'html.parser')
        table = soup.find('table', id='Table2')
        trs = table.find_all(
            'tr',
            class_='openMenu',
        )

        tr_batches = itertools.izip_longest(*[iter(trs)] * 2)

        documents = []
        for (tr0, tr1) in tr_batches:
            tds = tr0.find_all('td')

            document = Document(
                id_=int(urlparse.parse_qs(urlparse.urlparse(tr0['linkto']).query)['dilistid'][0]),
                checked='checked' in tds[1].input,
                source=tds[4].contents[2].string,
                source_no=u'{:s}字第{:d}號'.format(tds[5].string, int(tds[6].contents[1].string)),
                receiver=tds[9].contents[2].string,
                receive_datetime=datetime.datetime.strptime(tds[8].string.split(' ')[0], '%Y/%m/%d'),
                paper_nm=0,  # TODO
                subject=tr1.contents[1].contents[2].string.strip().split(u'：', 1)[1],
            )
            documents.append(document)

        return documents

    def receive_detail(self, document):
        params = {
            'menuCode': 'RECVQRY',
            'detail': 'showdetail',
            'dilistid': document.dilistid,
            'listid': '',
        }

        r = self.get('webeClient/main.php', params=params)
        soup = bs4.BeautifulSoup(r.content, 'html.parser')
        table = soup.find('table', id='Table1')
        trs = table.find_all('tr')

        tds = trs[1].find_all('td')
        document.paper_nm = tds[3].string

        as_ = trs[4].find_all('a')
        document.attachments = {}
        for a_ in as_[1:]:
            document.attachments[a_.string] = a_['href']

        return document


class Converter(object):
    @staticmethod
    def code(item):
        return u'{item.code_no} {item.code_nm}'.format(item=item)

    def __init__(self,
                 connection):

        self.connection = connection
        self.conductors = connection.select(
            from_='conductor',
            fields=[Connection.rtrim('user_nm')],
        )
        self.secrets = connection.select(
            from_='secret',
            fields=[
                Connection.rtrim('code_no'),
                Connection.rtrim('code_nm'),
            ],
        )
        self.papers = connection.select(
            from_='paper',
            fields=[
                Connection.rtrim('code_no'),
                Connection.rtrim('code_nm'),
            ],
        )
        self.books = connection.select(
            from_='book',
            fields=[
                Connection.rtrim('code_no'),
                Connection.rtrim('code_nm'),
            ],
        )

        secret_default = self.secrets[self.secrets.code_no == '1'].iloc[0]
        paper_default = self.papers[self.papers.code_no == '21'].iloc[0]
        book_default = self.books[self.books.code_no == '1'].iloc[0]

        self.archive_default = pd.DataFrame([{
            'secret': Converter.code(secret_default),
            'paper': Converter.code(paper_default),
            'book': Converter.code(book_default),
            'speed': u'普通件',
            'archive_no': 1,
            'process_days': 0.5,
            'measure': u'頁',
            'ymm_user': u'系統管理員',
        }])

    def to_archives(self, documents):
        now = datetime.datetime.now()

        archives = self.connection.select(
            from_='archive',
            fields=['sno', 'convert(int, receive_no)'],
            wheres=['Year(receive_date) = {:d}'.format(now.year)],
            order_bys=['sno desc'],
        )

        if archives.empty:
            sno = 1
            receive_no = 1
        else:
            sno = archives.sno.max() + 1
            receive_no = archives.receive_no.max() + 1

        new_archives = []
        for (num, document) in enumerate(documents):
            secret = self.secrets[self.secrets.code_nm == documents.secret_nm].iloc[0]
            paper = self.papers[self.papers.code_nm == documents.paper_nm].iloc[0]
            book = self.books[self.books.code_nm == documents.book_nm].iloc[0]

            archive = self.archive_default.copy()
            archive.update([{
                'ymm_year': now.year,
                'sno': sno + num,
                'receive_no': receive_no + num,
                'receive_date': '{:%Y/%m/%d}'.format(now),
                'secret': Converter.code(secret),
                'source': document.source,
                'source_no': document.source_no,
                'paper': Converter.code(paper),
                'book': Converter.code(book),
                'user_nm': document.user_nm,
                'subject': document.subject,
                'ymm_month': now.month,
            }])
            new_archives.append(archive)

        new_archives = pd.concat(new_archives, ignore_index=True)
        return new_archives


class Connection(pypyodbc.Connection):
    @staticmethod
    def rtrim(field):
        return 'rtrim(' + field + ')'

    @staticmethod
    def sentence(strs, sep=',', begin='', end='', default=''):
        begin = begin + ' '
        sep = ' ' + sep + ' '
        end = ' ' + end
        return (begin + sep.join(strs) + end) if strs else default

    def select(self,
               from_,
               top=None,
               fields=['*'],
               wheres=[],
               order_bys=[]):

        query = (
            u'select {top}{field} '
            u'from {from_} '
            u'{where} '
            u'{order_by} '
        ).format(
            top='' if top is None else 'top {} '.format(top),
            field=Connection.setence(fields),
            from_=from_,
            where=Connection.sentence(wheres, sep='and', begin='where'),
            order_by=Connection.sentence(order_bys, bein='order by'),
        )

        print(query)

        df = pd.read_sql(query, con=self)
        return df

    '''
    def insert(self,
               into,
               fields=[],
               values=[]):

        values = ['\'' + unicode(value) + '\'' for value in values]

        query = (
            u'insert '
            u'into {into} {field} '
            u'values {value} '
        ).format(
            into=into,
            field=Connection.sentence(fields, begin='(', end=')'),
            value=Connection.sentence(values, begin='(', end=')'),
        )

        print(query)
        cursor = self.cursor()
        cursor.execute(query)
        cursor.close()
        self.commit()
    '''
    def insert(self,
               df,
               into):

        df.to_sql(into, con=self, if_exists='append', index=False)
