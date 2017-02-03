# -*- coding: utf-8 -*-

import flask
import flask_bootstrap

from util import DEBUG, Document, Manager, eClient, Connection

app = flask.Flask(__name__)
app.debug = DEBUG
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
)


@app.route('/receive', methods=['GET', 'POST'])
def receive():
    messages = []
    alerts = []
    documents_by_source_no = manager.receive()

    if flask.request.method == 'POST':
        source_no = flask.request.form['source-no']
        conductor = flask.request.form['conductor']

        document = documents_by_source_no[source_no]
        document.user_nm = conductor
        manager.receive_detail(document)
        
        connection.insert(
            manager.to_archive(document), 
            into='archive',
        )

        manager.save_files(document)
        manager.set_checked(document)

        alerts.append(Manager.document_str(document, Manager.SUCCESS))

    return flask.render_template(
        'receive.html',
        messages=messages,
        alerts=alerts,
        documents=set(documents_by_source_no.values()),
        user_nms=manager.conductors.user_nm,
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=1235)