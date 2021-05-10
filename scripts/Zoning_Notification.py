import win32com.client, datetime, os, configparser, sys, traceback

try:
    # Assign start time variable for logging purposes
    StartTime = datetime.datetime.now().replace(microsecond=0)

    # Assign date variable for email subject line
    today = StartTime.today()

    # Assign config path
    root_path = os.path.dirname(os.path.dirname(__file__))
    ini_path = os.path.join(root_path, 'ini', 'notifications_config.ini')

    # Assign and read initialization file for required path information
    config = configparser.ConfigParser()
    config.read(ini_path)

    # Assign log object for outputting run-time details
    log_path = config.get('ZONING_INPUTS', 'Log_Path')
    log = open(log_path, "a")

    # Assign email you wish to have notified upon script completion
    email_recipient = config.get("ZONING_INPUTS", "Email_Recipient")
    email_cc = config.get("ZONING_INPUTS", "Email_CC")
    ini_month = config.get("GENERAL_INPUTS", "month")
    auto_month_trigger = config.get("GENERAL_INPUTS", "auto_month")

    # Assign today date variable automatically or through ini file based on auto month value
    if auto_month_trigger.lower() == 'true':
        month = datetime.datetime.today().strftime("%B")
    else:
        month = ini_month

    # Assign outlook object
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Build email to send to GIS Team if a new E-Designation file was detected
    print("Forming {} email".format(month))
    outlook = win32com.client.Dispatch("Outlook.Application")
    email_msg = outlook.CreateItem(0x0)
    email_msg.To = email_recipient
    email_msg.CC = email_cc
    email_msg.Subject = "Zoning Division monthly update reminder - {}".format(month)
    email_msg.HTMLBody = "<html><head></head><body>Greetings, <br />" \
               "<p>This is a monthly reminder to request information from the Zoning Division on all known zoning map " \
               "changes or text amendments that were adopted in <b>{}</b> that would produce changes to the following GIS files:</p>" \
               "<ul style='font-weight: bold';>" \
               "<li>DCP Initiated Rezonings</li>" \
               "<li>MIH (Appendix F)</li>" \
               "<li>Inclusionary Housing (VIH)</li>" \
               "<li>Transit Zones</li>" \
               "<li>Waterfront Access Plan (WAP)</li>" \
               "<li>Designated Areas in Manufacturing Districts (Appendix J)</li>" \
               "</ul>" \
               "<p>In addition, is the Zoning Division aware of any other mapped areas within the ZR that should " \
               "become new datasets?</p>" \
               "<p>Please list all changes and provide any supporting documentation. We appreciate your assistance.</p>" \
               "<p>Thank you and have a wonderful day!</p></body></html>".format(month)
    print("Email composition complete. Sending email.")
    email_msg.Send()
    print("Email sent")

    # Log total script run-time
    EndTime = datetime.datetime.now().replace(microsecond=0)
    print("Script runtime: {}".format(EndTime - StartTime))
    log.write(str(StartTime) + "\t" + str(EndTime) + "\t" + str(EndTime - StartTime) + "\n")
    log.close()

except Exception as e:
    print("error")
    print(e)
    tb = sys.exc_info()[2]
    tbinfo = traceback.format_tb(tb)[0]

    # Log any Python errors that were encountered during script run-time
    pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n" + str(sys.exc_info()[1])

    print(pymsg)

    log.write("" + pymsg + "\n")
    log.write("\n")
    log.close()