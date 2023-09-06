using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Uc;

namespace Teams.Tests
{
  internal static class Program
  {
    /// <summary>
    /// Der Haupteinstiegspunkt für die Anwendung.
    /// </summary>
    [STAThread]
    static void Main()
    {
      var version = "15.0.0.0";
      var teams = (IUCOfficeIntegration)new TeamsOfficeIntegration();

      var client = (IClient)teams.GetInterface(version, OIInterface.oiInterfaceILyncClient);

      var contactManager = client.ContactManager;

      // hier kontakt einfügen
      var uri = "felix@infomatik.eu";
      var contact = contactManager.GetContactByUri(uri);
      var availibility =
        (ContactAvailability)contact.GetContactInformation(ContactInformationType.ucPresenceAvailability);

      Console.WriteLine(availibility);
      Console.ReadLine();
    }

    [ComImport, Guid("00425F68-FFC1-445F-8EDF-EF78B84BA1C7")]
    public class TeamsOfficeIntegration
    {
    }
  }
}
