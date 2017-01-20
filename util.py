# -*- coding: utf-8 -*-

import bs4
import datetime
import itertools
import pandas as pd
import pypyodbc
import requests
import urlparse

pd.set_option('display.unicode.east_asian_width', True)

def now():
    return datetime.datetime.now()


class Document(object):
    def __init__(self,
                 eclient,
                 checked,
                 source,
                 source_word,
                 source_no,
                 date,
                 receiver,
                 subject,
                 dilistid):

        self.checked = checked
        self.eclient = eclient
        self.source = source
        self.source_word = source_word
        self.source_no = source_no
        self.date = date
        self.receiver = receiver
        self.subject = subject
        self.dilistid = dilistid

    def receive(self):
        params = {
            'menuCode': 'RECVQRY',
            'detail': 'showdetail',
            'dilistid': self.dilistid,
            'listid': '',
        }

        r = self.eclient.get('webeClient/main.php', params=params)
        soup = bs4.BeautifulSoup(r.content, 'html.parser')
        table = soup.find('table', id='Table1')
        trs = table.find_all('tr')

        tds = trs[1].find_all('td')
        self.paper = tds[3].string

        as_ = trs[4].find_all('a')
        self.attachments = {}
        for a_ in as_[1:]:
            self.attachments[a_.string] = a_['href']


class eClient(requests.Session):
    def __init__(self, 
                 server='http://localhost:80'):

        super(eClient, self).__init__()
        self.headers.update({'referer': ''})
        self.server = server

    def route(self, url):
        return urlparse.urljoin(self.server, url)
        
    def get(self, url, *args, **kwargs):
        return super(eClient, self).get(self.route(url), *args, **kwargs)

    def post(self, url, *args, **kwargs):
        return super(eClient, self).post(self.route(url), *args, **kwargs)

    def login(self, 
              userid='admin',
              passwd='admin_123'):

        self.post(
            'webeClient/menu.php',
            data={
                'login': '1',
                'userid': userid,
                'passwd': passwd,
            },
        )

    def receive(self, **kwargs):
        now_ = now()
        now_str = '{:d}-{:02d}-{:02d}'.format(
            now_.year - 1911,
            now_.month,
            now_.day,
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
                eclient=self,
                checked='checked' in tds[1].input,
                source=tds[4].contents[2].string,
                source_word=tds[5].string,
                source_no=int(tds[6].contents[1].string),
                date=datetime.datetime.strptime(tds[8].string.split(' ')[0], '%Y/%m/%d'),
                receiver=tds[9].contents[2].string,
                subject=tr1.contents[1].contents[2].string.strip().split(u'：', 1)[1],
                dilistid=int(urlparse.parse_qs(urlparse.urlparse(tr0['linkto']).query)['dilistid'][0]),
            )
            documents.append(document)

        return documents


class Config(object):
    def __init__(self):
        pass
        

class Connection(pypyodbc.Connection):
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
            field=', '.join(fields),
            from_=from_,
            where='where ' + ' and '.join(wheres) if wheres else '',
            order_by='order by ' + ', '.join(order_bys) if order_bys else '',
        )

        print(query)

        df = pd.read_sql(query, con=self, coerce_float=True)
        return df

    def insert(self,
               into,
               fields=[],
               values=[]):

        query = (
            u'insert '
            u'into {into} {field} '
            u'values {value} '
        ).format( 
            into=into,
            field='(' + ', '.join(fields) + ')' if fields else '',
            value='(' + ', '.join(map(u'\'{}\''.format, values)) + ')' if values else '',
        )

        print(query)
        '''
        cursor = self.cursor()
        cursor.execute(query)
        cursor.close()
        self.commit()
        '''