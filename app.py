def send_via_gmail(service, to, subject, body, is_html=False, attachment_path=None):
    message = MIMEMultipart("mixed")
    message["to"] = to
    message["subject"] = subject

    # Body
    alt_part = MIMEMultipart("alternative")
    if is_html:
        alt_part.attach(MIMEText(body, "html"))
    else:
        alt_part.attach(MIMEText(body, "plain"))
    message.attach(alt_part)

    # Attachment
    if attachment_path and os.path.exists(attachment_path):
        from email.mime.base import MIMEBase
        from email import encoders

        with open(attachment_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(attachment_path)}")
            message.attach(part)

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return service.users().messages().send(userId="me", body={"raw": raw}).execute()


@app.route("/send_bulk", methods=["POST"])
def send_bulk():
    creds = load_credentials()
    if not creds or not creds.valid:
        return redirect(url_for("authorize"))

    service = build("gmail", "v1", credentials=creds)

    recipients = extract_recipients(request.files.get("file"), request.form["recipients"])
    subjects = request.form["subjects"].splitlines()
    bodies = request.form["bodies"].split("===")
    pause = int(request.form["pause"])
    is_html = "html" in request.form

    # Save uploaded attachment
    attachment_file = request.files.get("attachment")
    attachment_path = None
    if attachment_file and attachment_file.filename:
        filename = secure_filename(attachment_file.filename)
        attachment_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        attachment_file.save(attachment_path)

    def background_task():
        for idx, recipient in enumerate(recipients, 1):
            subject = random.choice(subjects)
            body = random.choice(bodies)
            try:
                msg = send_via_gmail(
                    service,
                    recipient.strip(),
                    subject,
                    body.strip(),
                    is_html,
                    attachment_path
                )
                print(f"{idx}/{len(recipients)} Sent to {recipient} (ID: {msg['id']})")
            except Exception as e:
                print(f"{idx}/{len(recipients)} Failed {recipient}: {str(e)}")
            time.sleep(pause)

    threading.Thread(target=background_task).start()
    return f"Sending {len(recipients)} emails in background! You can close this page."
