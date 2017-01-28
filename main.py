# -*- coding: utf-8 -*-

import flask
import flask_bootstrap

from util import Document, Converter, eClient, Connection

app = flask.Flask(__name__)
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

converter = Converter(
    connection=connection,
)


@app.route('/receive', methods=['GET', 'POST'])
def receive():
    alerts = []

    documents = eclient.receive(
        # start_date='106-01-01',
    )

    if flask.request.method == 'POST':
        id_ = int(flask.request.form['id'])
        conductor = flask.request.form['conductor']

        new_documents = [
            eclient.receive_detail(document)
            for document in documents
            if document.id_ == id_
        ]
        new_documents[0].user_nm = conductor

        new_archives = converter.to_archives(new_documents)
        connection.insert(
            new_archives,
            into='archive',
        )
        ''' Set checked at eclient
        eclient.set_checked(document)
        '''

    return flask.render_template(
        'receive.html',
        alerts=alerts,
        documents=documents,
        user_nms=converter.conductors.user_nm,
    )

app.run(host='0.0.0.0', port=1234)
