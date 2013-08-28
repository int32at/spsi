/******************************************************************/
/*                                                                */
/*                SharePoint Solution Installer                   */
/*                                                                */
/*    Copyright 2007 Lars Fastrup Nielsen. All rights reserved.   */
/*    http://www.fastrup.dk                                       */
/*                                                                */
/*    This program contains the confidential trade secret         */
/*    information of Lars Fastrup Nielsen.  Use, disclosure, or   */
/*    copying without written consent is strictly prohibited.     */
/*                                                                */
/* KML: Added SiteCollectionFeatureId                             */
/* KML: Updated InstallOperation enum to be public                */
/* KML: Added BackWardCompatibilityConfigProps                    */
/* KML: Added ConfigProps                                         */
/* KML: Added RequireMoss, FeatureScope, SolutionId, FeatureId,   */
/*      SiteCollectionRelativeConfigLink, SSPRelativeConfigLink,  */
/*      DefaultDeployToSRP, and DocumentationUrl properties       */
/* KML: Added ConfigProps                                         */
/*                                                                */
/******************************************************************/
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;

using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System.Data.SqlClient;
using CodePlex.SharePointInstaller.Resources;


namespace CodePlex.SharePointInstaller
{
  internal class InstallConfiguration
  {
    #region Constants

    public class BackwardCompatibilityConfigProps
    {
      // "Apllication" mispelled on purpose to match original mispelling released
      public const string RequireDeploymentToCentralAdminWebApllication = "RequireDeploymentToCentralAdminWebApllication";
      // Require="MOSS" = RequireMoss="true" 
      public const string Require = "Require";
      // FarmFeatureId = FeatureId with FeatureScope = Farm
      public const string FarmFeatureId = "FarmFeatureId";
    }

    public class ConfigProps
    {
      public const string BannerImage = "BannerImage";
      public const string LogoImage = "LogoImage";
      public const string EULA = "EULA";
      public const string RequireMoss = "RequireMoss";
      public const string UpgradeDescription = "UpgradeDescription";
      public const string RequireDeploymentToCentralAdminWebApplication = "RequireDeploymentToCentralAdminWebApplication";
      public const string RequireDeploymentToAllContentWebApplications = "RequireDeploymentToAllContentWebApplications";
      public const string DefaultDeployToSRP = "DefaultDeployToSRP";
      public const string SolutionId = "SolutionId";
      public const string SolutionFile = "SolutionFile";
      public const string SolutionTitle = "SolutionTitle";
      public const string SolutionVersion = "SolutionVersion";
      public const string FeatureScope = "FeatureScope";
      public const string FeatureId = "FeatureId";
      public const string SiteCollectionRelativeConfigLink = "SiteCollectionRelativeConfigLink";
      public const string SSPRelativeConfigLink = "SSPRelativeConfigLink";
      public const string DocumentationUrl = "DocumentationUrl";
      public const string SupportedSharePointVersion = "SupportedSharePointVersion";
    }

    #endregion

    #region Internal Static Properties

    internal static string BannerImage
    {
      get { return ConfigurationManager.AppSettings[ConfigProps.BannerImage]; }
    }

    internal static string LogoImage
    {
      get { return ConfigurationManager.AppSettings[ConfigProps.LogoImage]; }
    }

    internal static string EULA
    {
      get { return ConfigurationManager.AppSettings[ConfigProps.EULA]; }
    }

    internal static bool RequireMoss
    {
      get
      {
        bool rtnValue = false;
        string valueStr = ConfigurationManager.AppSettings[ConfigProps.RequireMoss];
        if (String.IsNullOrEmpty(valueStr))
        {
          valueStr = ConfigurationManager.AppSettings[BackwardCompatibilityConfigProps.Require];
          rtnValue = valueStr != null && valueStr.Equals("MOSS", StringComparison.OrdinalIgnoreCase);
        }
        else
        {
          rtnValue = Boolean.Parse(valueStr);
        }
        return rtnValue;
      }
    }

    internal static Guid SolutionId
    {
      get { return new Guid(ConfigurationManager.AppSettings[ConfigProps.SolutionId]); }
    }

    internal static string SolutionFile
    {
      get { return ConfigurationManager.AppSettings[ConfigProps.SolutionFile]; }
    }

    internal static string SolutionTitle
    {
      get { return ConfigurationManager.AppSettings[ConfigProps.SolutionTitle]; }
    }

    internal static Version SolutionVersion
    {
      get { return new Version(ConfigurationManager.AppSettings[ConfigProps.SolutionVersion]); }
    }

    internal static string UpgradeDescription
    {
      get
      {
        string str = ConfigurationManager.AppSettings[ConfigProps.UpgradeDescription];
        if (str != null)
        {
          str = FormatString(str);
        }
        return str;
      }
    }

    internal static SPFeatureScope FeatureScope
    {
      get
      {
        // Default to farm features as this is what the installer only supported initially
        SPFeatureScope featureScope = SPFeatureScope.Farm;
        string valueStr = ConfigurationManager.AppSettings[ConfigProps.FeatureScope];
        if (!String.IsNullOrEmpty(valueStr))
        {
          featureScope = (SPFeatureScope)Enum.Parse(typeof(SPFeatureScope), valueStr, true);
        }
        return featureScope;
      }
    }

    // Modif JPI - Début
    internal static List<Guid?> FeatureId
    {
      get
      {
        string valueStr = ConfigurationManager.AppSettings[ConfigProps.FeatureId];

        //
        // Backwards compatibility with old configuration files before site collection features allowed
        //
        if (String.IsNullOrEmpty(valueStr))
        {
          valueStr = ConfigurationManager.AppSettings[BackwardCompatibilityConfigProps.FarmFeatureId];
        }

        if (!String.IsNullOrEmpty(valueStr))
        {
            string[] _strGuidArray = valueStr.Split(";".ToCharArray());
            if (_strGuidArray.Length >= 0)
            {
                List<Guid?> _guidArray = new List<Guid?>();
                foreach (string _strGuid in _strGuidArray)
                {
                    _guidArray.Add(new Guid(_strGuid));
                }
                return _guidArray;
            }
        }

        return null;
      }
    }
    // Modif JPI - Fin

    internal static bool RequireDeploymentToCentralAdminWebApplication
    {
      get
      {
        string valueStr = ConfigurationManager.AppSettings[ConfigProps.RequireDeploymentToCentralAdminWebApplication];

        //
        // Backwards compatability with old configuration files containing spelling error in the 
        // application setting key (Bug 990).
        //
        if (String.IsNullOrEmpty(valueStr))
        {
          valueStr = ConfigurationManager.AppSettings[BackwardCompatibilityConfigProps.RequireDeploymentToCentralAdminWebApllication];
        }

        if (!String.IsNullOrEmpty(valueStr))
        {
          return valueStr.Equals("true", StringComparison.OrdinalIgnoreCase);
        }

        return false;
      }
    }

    internal static bool RequireDeploymentToAllContentWebApplications
    {
      get
      {
        string valueStr = ConfigurationManager.AppSettings[ConfigProps.RequireDeploymentToAllContentWebApplications];
        if (!String.IsNullOrEmpty(valueStr))
        {
          return valueStr.Equals("true", StringComparison.OrdinalIgnoreCase);
        }
        return false;
      }
    }

    internal static bool DefaultDeployToSRP
    {
      get
      {
        string valueStr = ConfigurationManager.AppSettings[ConfigProps.DefaultDeployToSRP];
        if (!String.IsNullOrEmpty(valueStr))
        {
          return valueStr.Equals("true", StringComparison.OrdinalIgnoreCase);
        }
        return false;
      }
    }

    internal static Version InstalledVersion
    {
      get
      {
        try
        {
          SPFarm farm = SPFarm.Local;
          string key = "Solution_" + SolutionId.ToString() + "_Version";
          return farm.Properties[key] as Version;
        }

        catch (NullReferenceException ex)
        {
          throw new InstallException(CommonUIStrings.installExceptionDatabase, ex);
        }

        catch (SqlException ex)
        {
          throw new InstallException(ex.Message, ex);
        }
      }

      set
      {
        try
        {
          SPFarm farm = SPFarm.Local;
          string key = "Solution_" + SolutionId.ToString() + "_Version";
          farm.Properties[key] = value;
          farm.Update();
        }

        catch (NullReferenceException ex)
        {
            throw new InstallException(CommonUIStrings.installExceptionDatabase, ex);
        }

        catch (SqlException ex)
        {
          throw new InstallException(ex.Message, ex);
        }
      }
    }

    public static bool ShowFinishedControl
    {
      get
      {
        return !String.IsNullOrEmpty(ConfigurationManager.AppSettings[ConfigProps.SiteCollectionRelativeConfigLink]) ||
          !String.IsNullOrEmpty(ConfigurationManager.AppSettings[ConfigProps.SSPRelativeConfigLink]);
      }
    }

    public static string SiteCollectionRelativeConfigLink
    {
      get { return ConfigurationManager.AppSettings[ConfigProps.SiteCollectionRelativeConfigLink]; }
    }

    public static string SSPRelativeConfigLink
    {
      get { return ConfigurationManager.AppSettings[ConfigProps.SSPRelativeConfigLink]; }
    }

    public static string DocumentationUrl
    {
      get { return ConfigurationManager.AppSettings[ConfigProps.DocumentationUrl]; }
    }

    public static string SupportedSharePointVersion
    {
        get { return ConfigurationManager.AppSettings[ConfigProps.SupportedSharePointVersion]; }
    }
    #endregion

    #region Internal Static Methods

    internal static string FormatString(string str)
    {
      return FormatString(str, null);
    }

    internal static string FormatString(string str, params object[] args)
    {
      string formattedStr = str;
      string solutionTitle = SolutionTitle;
      if (!String.IsNullOrEmpty(solutionTitle))
      {
        formattedStr = formattedStr.Replace("{SolutionTitle}", solutionTitle);
      }
      if (args != null)
      {
        formattedStr = String.Format(formattedStr, args);
      }
      return formattedStr;
    }

    #endregion
  }

  public enum InstallOperation
  {
    Install,
    Upgrade,
    Repair,
    Uninstall
  }
}
