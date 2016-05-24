from flask import Flask, request, redirect, url_for, render_template, send_file
import os
import json
import glob
import scrapermod
from uuid import uuid4
from rq import Queue
from rq.job import Job
from worker import conn

q = Queue(connection=conn)

app = Flask(__name__)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    """Handle the upload of a file."""
    form = request.form

    # Create a unique "session ID" for this particular batch of uploads.
    upload_key = str(uuid4())

    # Is the upload using Ajax, or a direct POST by the form?
    is_ajax = False
    if form.get("__ajax", None) == "true":
        is_ajax = True

    # Target folder for these uploads.
    target = "uploadr/static/uploads"

    print "=== Form Data ==="
    for key, value in form.items():
        print key, "=>", value

    for upload in request.files.getlist("file"):
        filename = upload.filename.rsplit("/")[0]
        filename = "tempuploadfile.xlsx"
        destination = "/".join([target, filename])
        print "Accept incoming file:", filename
        print "Save it to:", destination
        upload.save(destination)

    if is_ajax:
        return ajax_response(True, upload_key)
    else:
        return redirect(url_for("upload_complete", uuid=upload_key))


@app.route("/files/<uuid>")
def upload_complete(uuid):
    """The location we send them to at the end of the upload."""

    # Get their files.
    root = "uploadr/static/uploads"
    if not os.path.isdir(root):
        return "Error: UUID not found!"

    files = []
    for file in glob.glob("{}/*.*".format(root)):
        fname = file.split(os.sep)[-1]
        files.append(fname)
        fpath = root + '/' + fname
        job = q.enqueue_call(
            func=scrapermod.scrape_from_pdf, args=(fpath,), result_ttl=5000
        )
        print "job_id: %s" % job.get_id()
        #print returnfile

    #return returnfile
    return render_template('files.html', filesjobs=zip([fpath],[str(job.get_id())]) )
    #return send_file('../' + returnfile, attachment_filename=returnfile)

@app.route("/results/<job_key>", methods=['GET'])
def get_results(job_key):

    job = Job.fetch(job_key, connection=conn)

    if job.is_finished:
        return send_file('../' + job.result, attachment_filename=job.result), 200
    else:
        return "Nay!", 202


def ajax_response(status, msg):
    status_code = "ok" if status else "error"
    return json.dumps(dict(
        status=status_code,
        msg=msg,
    ))
