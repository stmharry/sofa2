# -*- coding: utf-8 -*-

import bs4 
import flask
import flask_bootstrap
import util

app = flask.Flask(__name__)
flask_bootstrap.Bootstrap(app)

eclient = util.Eclient(server='http://10.167.131.227')
eclient.login()

@app.route('/receive')
def receive():
    documents = eclient.receive(
        startDate='106-01-01',
        endDate='106-01-10',
    )

    return flask.render_template('receive.html', documents=documents)

app.run(host='0.0.0.0', port=1234)