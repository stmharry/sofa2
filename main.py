# -*- coding: utf-8 -*-

import bs4 
import flask
import flask_bootstrap
import util

app = flask.Flask(__name__)
flask_bootstrap.Bootstrap(app)

eclient = util.Eclient(server='http://10.167.131.227')
eclient.login()

connection = util.Connection(
    driver=u'{SQL Server Native Client 11.0}',
    server=u'10.167.131.227,1433',
    database=u'YMM_POLICE',
    uid=u'sa',
    pwd=u'twntfs@cloud',
    unicode_results=True,
)

@app.route('/receive', methods=['GET', 'POST'])
def receive():
    alerts = []
    if flask.request.method == 'POST':
        alerts.append(u'LOG: dilistid={}, 承辦人={}'.format(
            flask.request.form['dilistid'], 
            flask.request.form['conductor'], 
        ))

    documents = eclient.receive(
        #startDate='106-01-01',
    )
    conductors = connection.select(
        from_='conductor', 
        fields=['user_nm'],
    )

    return flask.render_template(
        'receive.html',
        alerts=alerts,
        documents=documents,
        conductors=conductors,
    )


app.run(host='0.0.0.0', port=1234)