# -*- coding: utf-8 -*-

import flask
import flask_bootstrap

from util import Manager, eClient, Connection

app = flask.Flask(__name__)
app.debug = True
flask_bootstrap.Bootstrap(app)

eclient = eClient(
    server='http://10.167.131.227',
    userid='admin',
    passwd='admin_123',
)

connection = Connection(
    driver='{SQL Server Native Client 11.0}',
    server='10.167.131.227,1433',
    database='YMM_POLICE',
    uid='sa',
    pwd='twntfs@cloud',
    unicode_results=True,
)

manager = Manager(
    eclient=eclient,
    connection=connection,
    config_path='config.cfg',
)


@app.route('/receive', methods=['GET', 'POST'])
def receive():
    eclient.login()

    manager.debug_messages = []
    manager.alerts = []
    
    document = None
    document_by_source_no = manager.receive()

    if flask.request.method == 'POST':
        source_no = flask.request.form['source-no']
        conductor = flask.request.form['conductor']

        document = document_by_source_no[source_no]
        manager.process(document, conductor=conductor)

    return flask.render_template(
        'receive.html',
        manager=manager,
        document=document,
        documents=document_by_source_no.values(),
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=1235)
