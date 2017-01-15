# -*- coding: utf-8 -*-

import bs4 
import flask
import util

app = flask.Flask(__name__)
eclient = util.Eclient(server='http://10.167.131.227')
eclient.login()

@app.route('/rcv')
def receive():
    soup = bs4.BeautifulSoup('', 'html.parser')
    soup.append(soup.new_tag('ol'))

    documents = eclient.receive(
        startDate='106-01-01',
        endDate='106-01-10',
    )

    for document in documents:
        li = soup.new_tag('li')
        #import pdb; pdb.set_trace()
        li.string = unicode(document)
        soup.ol.append(li)

    eclient.receive_detail(documents[0])

    #print(unicode(soup))
    return flask.make_response(str(soup))

app.run(host='0.0.0.0', port=1234)
#receive()   