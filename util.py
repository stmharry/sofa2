# -*- coding: utf-8 -*-

import bs4
import datetime
import itertools
import pypyodbc
import requests
import urlparse

class Document(object):
    def __str__(self):
        if self.attachments is not None:
            attachment_str = u'附件: {}\n'.format(', '.join(self.attachments.keys()))
        else:
            attachment_str = u''

        return (
            u'{document.date:%Y-%m-%d}'
            u'{document.source_word}字第{document.source_no}號\n'
            u'{document.subject}\n'
            u'{attachment_str}\n'
        ).format(
            document=self,
            attachment_str=attachment_str,
        )

    def __repr__(self):
        return (
            u'{document.date:%Y-%m-%d} {document.source_word}字第{document.source_no}號'
        ).format(document=self)

    def __init__(self,
                 source_name=None,
                 source_word=None,
                 source_no=None,
                 date=None,
                 subject=None,
                 dilistid=None,
                 attachments=None,
                 **kwargs):

        self.source_name = source_name
        self.source_word = source_word
        self.source_no = source_no
        self.date = date
        self.subject = subject
        self.dilistid = dilistid
        self.attachments = attachments
        self.__dict__.update(kwargs)


class Eclient(object):
    def __init__(self, 
                 server='http://localhost:80'):

        self.sess = requests.Session()
        self.sess.headers.update({'referer': ''})
        self.server = server

    def route(self, url):
        return urlparse.urljoin(self.server, url)

    def login(self, 
              userid='admin',
              passwd='admin_123'):

        self.sess.post(
            self.route('webeClient/menu.php'),
            data={
                'login': '1',
                'userid': userid,
                'passwd': passwd,
            },
        )

    def receive(self, **kwargs):
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
            'doc_word': '',
            'doc_no': '',
            'doc_title': '',
            'full_compare': '0',
            'senderid': '',
            'sendername': '',
            'query_time': 'create_time',
            'startDate': now_str,
            'sHour': '00',
            'endDate': now_str,
            'eHour': '23',
            'noteonly': '2',
            'pn': '0',
            'ifquery_recvqry': '1',
            'full_compare_send': '',
        }
        params.update(kwargs)
        r = self.sess.get(self.route('webeClient/main.php'), params=params)

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
                source_name=tds[4].contents[2].string,
                source_word=tds[5].string,
                source_no=int(tds[6].contents[1].string),
                date=datetime.datetime.strptime(tds[8].string.split(' ')[0], '%Y/%m/%d'),
                subject=tr1.contents[1].contents[2].string.strip().split(u'：', 1)[1],
                dilistid=int(urlparse.parse_qs(urlparse.urlparse(tr0['linkto']).query)['dilistid'][0]),
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

        r = self.sess.get(self.route('webeClient/main.php'), params=params)
        soup = bs4.BeautifulSoup(r.content, 'html.parser')
        as_ = soup.select('table#Table1 a')

        document.attachments = {}
        for a_ in as_[1:]:
            document.attachments[a_.string] = a_['href']


class SQLServer(object):
    def __init__(self, 
                 driver=u'{SQL Server Native Client 11.0}',
                 server=u'隊本部收發\SQLEXPRESS',
                 database=u'YMM_POLICE',
                 uid=u'sa',
                 pwd=u'twntfs@cloud'):

    self.conn = pypyodbc.connect(
        driver=driver,
        server=server,
        database=database,
        uid=uid,
        pwd=pwd,
    )

    cur = conn.cursor()

        cur.execute((
            'select top 10 * '
            'from dbo.archive as x '
            'where x.secret = ?'
        ), (u'1 普通',))

        for row in cur.fetchall():
            print(row['receive_no'])