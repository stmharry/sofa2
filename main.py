# -*- coding: utf-8 -*-

import bs4 
import flask
import flask_bootstrap
import pandas as pd
import util

app = flask.Flask(__name__)
flask_bootstrap.Bootstrap(app)

eclient = util.eClient(server='http://10.167.131.227')
eclient.login()

connection = util.Connection(
    driver=u'{SQL Server Native Client 11.0}',
    server=u'10.167.131.227,1433',
    database=u'YMM_POLICE',
    uid=u'sa',
    pwd=u'twntfs@cloud',
    unicode_results=True,
)

conductors = connection.select(
    from_='conductor', 
    fields=['user_nm'],
)

secrets = connection.select(
    from_='secret',
    fields=['code_no', 'code_nm'],
)
secret_default = secrets[secrets.code_no.map(unicode.strip) == '1'].loc[0]

books = connection.select(
    from_='book',
    fields=['code_no', 'code_nm'],
)
book_default = books[books.code_no.map(unicode.strip) == '1'].loc[0]

speed_default = u'普通件'
archive_no_default = 1
process_days_default = 0.5
measure_default = u'頁'
user_default = u'系統管理員'

@app.route('/receive', methods=['GET', 'POST'])
def receive():
    alerts = []

    documents = eclient.receive(
        #startDate='106-01-01',
    )

    if flask.request.method == 'POST':
        dilistid = int(flask.request.form['dilistid'])
        conductor = flask.request.form['conductor']

        document = [document for document in documents if document.dilistid == dilistid]
        assert len(document) == 1
        document = document[0]
        document.receive()
    
        archives = connection.select(
            from_='archive', 
            #fields=['sno', 'receive_no'],
            wheres=['Year(receive_date) = {}'.format(util.now().year)],
            order_bys=['sno desc'],
        )
        alerts.append(archives.iloc[0])

        now = util.now()
        field_values = {
            'ymm_year': now.year, 
            'sno': max(archives.sno.tolist() or [0]) + 1,
            'receive_no': max(archives.receive_no.map(int).tolist() or [0]) + 1,
            'receive_date': '{:%Y/%m/%d}'.format(now),
            'secret': u'{secret.code_no}{secret.code_nm}'.format(secret=secret_default),
            'source': document.source,
            'source_no': u'{document.source_word}字第{document.source_no}號'.format(document=document),
            'paper': document.paper,
            'book': u'{book.code_no}{book.code_nm}'.format(book=book_default),
            'speed': speed_default,
            'user_nm': conductor,
            'subject': document.subject,
            'archive_no': archive_no_default,
            'process_days': process_days_default,
            'measure': measure_default,
            'ymm_month': '{:02d}'.format(now.month),
            'ymm_user': user_default,
        }

        connection.insert(
            into='archive',
            fields=field_values.keys(),
            values=field_values.values(),
        )

    return flask.render_template(
        'receive.html',
        alerts=alerts,
        documents=documents,
        user_nms=conductors.user_nm,
    )


app.run(host='0.0.0.0', port=1234)