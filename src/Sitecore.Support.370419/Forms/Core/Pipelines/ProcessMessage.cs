// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ProcessMessage.cs" company="Sitecore A/S">
//   Copyright (C) 2010 by Sitecore A/S
// </copyright>
// <summary>
// Process Message  
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Sitecore.Support.Forms.Core.Pipelines
{
  using System;
  using System.Net.Mail;
  using System.Text.RegularExpressions;
  using Sitecore.Data;
  using Sitecore.Data.Items;
  using Sitecore.Diagnostics;
  using Sitecore.Links;
  using Sitecore.StringExtensions;
  using Sitecore.WFFM.Abstractions.Actions;
  using Sitecore.WFFM.Abstractions.Dependencies;
  using Sitecore.WFFM.Abstractions.Mail;
  using Sitecore.WFFM.Abstractions.Shared;
  using Sitecore.WFFM.Abstractions.Utils;

  /// <summary>
  /// Processors for Process message
  /// </summary>
  public class ProcessMessage
  {
    /// <summary>
    /// External dependency for initialization from processors
    /// </summary>
    public IItemRepository ItemRepository { get; set; }
    public IFieldProvider FieldProvider { get; set; }

    public ProcessMessage()
      : this(DependenciesManager.WebUtil)
    {
    }

    public ProcessMessage(IWebUtil webUtil)
    {
      Assert.IsNotNull(webUtil, "webUtil");
      srcReplacer = string.Join(string.Empty, new[] { "src=\"", webUtil.GetServerUrl(), "/~" });
      shortHrefReplacer = string.Join(string.Empty, new[] { "href=\"", webUtil.GetServerUrl(), "/" });
      shortHrefMediaReplacer = string.Join(string.Empty, new[] { "href=\"", webUtil.GetServerUrl(), "/~/" });
      hrefReplacer = shortHrefReplacer + "~";
    }

    private readonly string srcReplacer;
    private readonly string shortHrefReplacer;
    private readonly string shortHrefMediaReplacer;
    private readonly string hrefReplacer;

    public void ExpandLinks(ProcessMessageArgs args)
    {
      var mail = LinkManager.ExpandDynamicLinks(args.Mail.ToString());
      args.Mail.Remove(0, args.Mail.Length);
      args.Mail.Append(mail);
    }

    public void ExpandTokens(ProcessMessageArgs args)
    {
      Assert.IsNotNull(this.ItemRepository, "ItemRepository");
      Assert.IsNotNull(this.FieldProvider, "FieldProvider");
      foreach (AdaptedControlResult field in args.Fields)
      {
        var item = this.ItemRepository.CreateFieldItem(this.ItemRepository.GetItem(field.FieldID));

        string value = field.Value;
        value = this.FieldProvider.GetAdaptedValue(field.FieldID, value);
        value = Regex.Replace(value, "src=\"/sitecore/shell/themes/standard/~", this.srcReplacer);
        value = Regex.Replace(value, "href=\"/sitecore/shell/themes/standard/~", this.hrefReplacer);
        value = Regex.Replace(value, "on\\w*=\".*?\"", string.Empty);

        if (args.MessageType == MessageType.Sms)
        {
          args.Mail.Replace("[{0}]".FormatWith(item.FieldDisplayName), value);
          args.Mail.Replace("[{0}]".FormatWith(item.Name), value);
        }
        else
        {
          if (!string.IsNullOrEmpty(field.Parameters) && args.IsBodyHtml)
          {
            if (field.Parameters.StartsWith("multipleline"))
            {
              value = value.Replace(Environment.NewLine, "<br/>");
            }
            if (field.Parameters.StartsWith("secure") && field.Parameters.Contains("<schidden>"))
            {
              value = Regex.Replace(value, @"\d", "*");
            }
          }

          //#405654
          var replaced = args.Mail.ToString();
          if (Regex.IsMatch(replaced, "\\[<label id=\"" + item.ID + "\">[^<]+?</label>]"))
          {
            replaced = Regex.Replace(replaced, "\\[<label id=\"" + item.ID + "\">[^<]+?</label>]", value);
          }
          if (Regex.IsMatch(replaced, "\\[<label id=\"" + item.ID + "\" renderfield=\"Value\">[^<]+?</label>]"))
          {
            replaced = Regex.Replace(replaced, "\\[<label id=\"" + item.ID + "\" renderfield=\"Value\">[^<]+?</label>]", field.Value);
          }
          if (Regex.IsMatch(replaced, "\\[<label id=\"" + item.ID + "\" renderfield=\"Text\">[^<]+?</label>]"))
          {
            replaced = Regex.Replace(replaced, "\\[<label id=\"" + item.ID + "\" renderfield=\"Text\">[^<]+?</label>]", value);
          }
          replaced = replaced.Replace(item.ID.ToString(), value);
          args.Mail.Clear().Append(replaced);
        }


        args.From = args.From.Replace("[" + item.ID + "]", value);
        args.From = args.From.Replace(item.ID.ToString(), value);
        args.To.Replace(string.Join(string.Empty, new[] { "[", item.ID.ToString(), "]" }), value);
        args.To.Replace(string.Join(string.Empty, new[] { item.ID.ToString() }), value);
        args.CC.Replace(string.Join(string.Empty, new[] { "[", item.ID.ToString(), "]" }), value);
        args.CC.Replace(string.Join(string.Empty, new[] { item.ID.ToString() }), value);
        args.Subject.Replace(string.Join(string.Empty, new[] { "[", item.ID.ToString(), "]" }), value);

        #region fix
        args.From = args.From.Replace("[Value: " + item.FieldDisplayName + "]", field.Value);
        args.To.Replace(string.Join(string.Empty, new[] { "[Value: ", item.FieldDisplayName, "]" }), field.Value);
        args.CC.Replace(string.Join(string.Empty, new[] { "[Value: ", item.FieldDisplayName, "]" }), field.Value);
        args.Subject.Replace(string.Join(string.Empty, new[] { "[Value: ", item.FieldDisplayName, "]" }), field.Value);

        args.From = args.From.Replace("[Text: " + item.FieldDisplayName + "]", value);
        args.To.Replace(string.Join(string.Empty, new[] { "[Text: ", item.FieldDisplayName, "]" }), value);
        args.CC.Replace(string.Join(string.Empty, new[] { "[Text: ", item.FieldDisplayName, "]" }), value);
        args.Subject.Replace(string.Join(string.Empty, new[] { "[Text: ", item.FieldDisplayName, "]" }), value);
        #endregion

        args.From = args.From.Replace("[" + item.FieldDisplayName + "]", value);
        args.To.Replace(string.Join(string.Empty, new[] { "[", item.FieldDisplayName, "]" }), value);
        args.CC.Replace(string.Join(string.Empty, new[] { "[", item.FieldDisplayName, "]" }), value);
        args.Subject.Replace(string.Join(string.Empty, new[] { "[", item.FieldDisplayName, "]" }), value);

        args.From = args.From.Replace("[" + field.FieldName + "]", value);
        args.To.Replace(string.Join(string.Empty, new[] { "[", field.FieldName, "]" }), value);
        args.CC.Replace(string.Join(string.Empty, new[] { "[", field.FieldName, "]" }), value);
        args.Subject.Replace(string.Join(string.Empty, new[] { "[", field.FieldName, "]" }), value);
      }
    }

    public void AddHostToItemLink(ProcessMessageArgs args)
    {
      args.Mail.Replace("href=\"/", this.shortHrefReplacer);
    }

    public void AddHostToMediaItem(ProcessMessageArgs args)
    {
      args.Mail.Replace("href=\"~/", this.shortHrefMediaReplacer);
    }

    public void AddAttachments(ProcessMessageArgs args)
    {
      Assert.IsNotNull(this.ItemRepository, "ItemRepository");

      if (!args.IncludeAttachment)
      {
        return;
      }

      foreach (AdaptedControlResult field in args.Fields)
      {
        if (!string.IsNullOrEmpty(field.Parameters) && field.Parameters.StartsWith("medialink") && !string.IsNullOrEmpty(field.Value))
        {
          var uri = ItemUri.Parse(field.Value);
          if (uri != null)
          {
            var item = this.ItemRepository.GetItem(uri);
            if (item != null)
            {
              var mediItem = new MediaItem(item);
              args.Attachments.Add(new Attachment(mediItem.GetMediaStream(),
                string.Join(".", new[]
                {
                  mediItem.Name, mediItem.Extension
                }),
                mediItem.MimeType));
            }
          }
        }
      }
    }

    public void BuildToFromRecipient(ProcessMessageArgs args)
    {
      if (!string.IsNullOrEmpty(args.Recipient) && !string.IsNullOrEmpty(args.RecipientGateway))
      {
        if (args.To.Length > 0)
        {
          args.To.Remove(0, args.To.Length);
        }
        args.To.Append(args.Fields.GetValueByFieldID(args.Recipient)).Append(args.RecipientGateway);
      }
    }


    public void SendEmail(ProcessMessageArgs args)
    {
      var client = new SmtpClient(args.Host)
      {
        EnableSsl = args.EnableSsl
      };

      if (args.Port != 0)
      {
        client.Port = args.Port;
      }

      client.Credentials = args.Credentials;
      client.Send(this.GetMail(args));
    }

    private MailMessage GetMail(ProcessMessageArgs args)
    {
      if (!Regex.Match(args.From, Sitecore.Constants.EmailRegExp).Success)
      {
        throw new Exception("The email message was not sent.The email address specified in the \"From\" parameter is not valid.");
      }

      var mail = new MailMessage(args.From.Replace(";", ","), args.To.Replace(";", ",").ToString(),
        args.Subject.ToString(), args.Mail.ToString()) { IsBodyHtml = args.IsBodyHtml };

      if (args.CC.Length > 0)
      {
          char[] separator = new char[] { ',' };
          foreach (string str in args.CC.Replace(";", ",").ToString().Split(separator))
          {
              mail.CC.Add(new MailAddress(str));
          }
      }
      if (args.BCC.Length > 0)
      {
          char[] chArray2 = new char[] { ',' };
          foreach (string str2 in args.BCC.Replace(";", ",").ToString().Split(chArray2))
          {
              mail.Bcc.Add(new MailAddress(str2));
          }
      }
      args.Attachments.ForEach(attachment => mail.Attachments.Add(attachment));

      return mail;
    }
  }
}
