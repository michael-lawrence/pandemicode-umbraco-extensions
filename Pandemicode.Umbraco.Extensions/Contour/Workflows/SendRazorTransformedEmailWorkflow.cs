using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using System.Web;
using System.Xml;
using RazorEngine;
using umbraco;
using Umbraco.Forms.Core;
using Umbraco.Forms.Core.Attributes;
using Umbraco.Forms.Core.Enums;
using Umbraco.Forms.Data.Storage;
using System.Linq;
using umbraco.MacroEngines;
using umbraco.BusinessLogic;

namespace Pandemicode.Umbraco.Extensions.Workflows {
	public class SendRazorTransformedEmailWorkflow : WorkflowType {
		[Setting("Email", description = "Enter the receiver email", control = "Umbraco.Forms.Core.FieldSetting.TextField")]
		public string Email { get; set; }

		[Setting("Subject", description = "Enter the subject", control = "Umbraco.Forms.Core.FieldSetting.TextField")]
		public string Subject { get; set; }

		[Setting("RazorFile", description = "The .cshtml file to transform the record. (<a href='/umbraco/plugins/pandemicode/razor/sendRazorEmailSample.cshtml.txt' target='_blank'>Sample File</a>)", control = "Umbraco.Forms.Core.FieldSetting.File")]
		public string RazorFile { get; set; }

		public static int pageID = -1;

		public SendRazorTransformedEmailWorkflow() {
			this.Name = "Send Razor Transformed Email";
			this.Id = new Guid("C68353CC-EC1F-4E80-A3B5-01FD27FA9AF3");
			this.Description = "Send the result of the form to an email address";
		}

		public override WorkflowExecutionStatus Execute(Record record, RecordEventArgs e) {
			try {
				MailMessage m = new MailMessage();
				m.From = new MailAddress(UmbracoSettings.NotificationEmailSender);
				m.Subject = Subject;
				m.IsBodyHtml = true;

				if (this.Email.Contains(";")) {
					string[] emails = this.Email.Split(';');
					foreach (string email in emails) {
						m.To.Add(email.Trim());
					}
				} else {
					m.To.Add(this.Email);
				}

				RecordsViewer viewer = new RecordsViewer();
				XmlNode xml = viewer.GetSingleXmlRecord(record, new XmlDocument());

				string result = xml.OuterXml;

				if (!string.IsNullOrEmpty(this.RazorFile)) {
					string fullPath = HttpContext.Current.Request.MapPath(this.RazorFile);
					using (StreamReader sr = File.OpenText(fullPath)) {
						string template = sr.ReadToEnd();
						result = Razor.Parse(template, new { Record = record });
					}
				}

				m.Body = result;

				

				Log.Add(LogTypes.Debug, record.UmbracoPageId, result);

				try {
					SmtpClient s = new SmtpClient();
					s.Send(m);
				} catch (SmtpException ex) {
					Log.Add(LogTypes.Error, record.UmbracoPageId, ex.ToString());
				} catch (Exception ex) {
					Log.Add(LogTypes.Error, record.UmbracoPageId, ex.ToString());
				}
			} catch (Exception ex) {
				Log.Add(LogTypes.Error, record.UmbracoPageId, ex.ToString());
			}

			return WorkflowExecutionStatus.Completed;
		}

		public override List<Exception> ValidateSettings() {
			List<Exception> l = new List<Exception>();

			if (String.IsNullOrEmpty(this.Email))
				l.Add(new Exception("'Email' setting not filled out"));

			if (String.IsNullOrEmpty(this.Subject))
				l.Add(new Exception("'Subject' setting not filled out"));

			return l;
		}
	}
}