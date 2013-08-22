/**********************************************************************/
/*                                                                    */
/*                   SharePoint Solution Installer                    */
/*             http://www.codeplex.com/sharepointinstaller            */
/*                                                                    */
/*               (c) Copyright 2007 Lars Fastrup Nielsen.             */
/*                                                                    */
/*  This source is subject to the Microsoft Permissive License.       */
/*  http://www.codeplex.com/sharepointinstaller/Project/License.aspx  */
/*                                                                    */
/**********************************************************************/
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;
using System.Windows.Forms;

namespace CodePlex.SharePointInstaller
{
  public class InstallerControl : UserControl
  {
    private string title;
    private string subTitle;

    protected InstallerControl()
    {
    }

    public string Title
    {
      get { return title; }
      set { title = value; }
    }

    public string SubTitle
    {
      get { return subTitle; }
      set { subTitle = value; }
    }

    protected InstallerForm Form
    {
      get
      {
        return (InstallerForm)this.ParentForm;
      }
    }

    protected internal virtual void Open(InstallOptions options) {}

    protected internal virtual void Close(InstallOptions options) {}

    protected internal virtual void RequestCancel() 
    {
      Application.Exit();
    }

      private void InitializeComponent()
      {
          this.SuspendLayout();
          // 
          // InstallerControl
          // 
          this.Name = "InstallerControl";
          this.ResumeLayout(false);

      }

  }

  public class InstallerControlList : List<InstallerControl>
  {
  }
}
