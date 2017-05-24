// -----------------------------------------------------------------------------------------------------------
// <copyright file="Program.cs">
// Copyright © 2017 George Chakhidze. All rights reserved.
// </copyright>
// <author>George Chakhidze</author>
// <email>0xfeeddeadbeef@gmail.com</email>
// <date>Monday, May 22, 2017 5:04:30 PM</date>
// <summary>Connects to LibreOffice Server and converts DOCX to PDF.</summary>
// -----------------------------------------------------------------------------------------------------------

using System;
using System.IO;
using System.Net.Sockets;
using Microsoft.Win32;
using uno;
using uno.util;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.bridge;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.util;

internal static class Program
{
    //
    // PREREQUISITES:
    //   - Install LibreOffice 64 bit (https://www.libreoffice.org/)
    //   - Run in server mode (see Launch-LibreOfficeServer.ps1 script)
    //

    private static void Main(string[] args)
    {
        ConvertToPDF(
            serverHost: "localhost",
            serverPort: 2002,
            inputFile:  new Uri(@"C:\Test\Word.docx"),
            outputFile: new Uri(@"C:\Test\Adobe.pdf")
        );
    }

    private static void ConvertToPDF(string serverHost, int serverPort, Uri inputFile, Uri outputFile)
    {
        string libreOfficePath = GetLibreOfficePath();

        if (string.IsNullOrEmpty(libreOfficePath) || !Directory.Exists(libreOfficePath))
        {
            throw new InvalidOperationException("LibreOffice is not installed.");
        }

        // HACK: This is much more convenient than cluttering system-wide environment.
        //       These variables must be set to correct paths, or else, this method will fail (!!!)
        Environment.SetEnvironmentVariable("UNO_PATH", libreOfficePath, EnvironmentVariableTarget.Process);
        Environment.SetEnvironmentVariable("PATH", Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Process) + ";" + libreOfficePath, EnvironmentVariableTarget.Process);


        // FIX: Workaround for a known bug: XUnoUrlResolver forgets to call WSAStartup. We can use dummy Socket for that:
        using (var dummySocket = new Socket(AddressFamily.InterNetwork, SocketType.Raw, ProtocolType.IP))
        {
            // First, initialize local service manager (IoC container of some kind):
            XComponentContext componentContext = Bootstrap.bootstrap();
            XMultiComponentFactory localServiceManager = componentContext.getServiceManager();
            XUnoUrlResolver urlResolver = (XUnoUrlResolver)localServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", componentContext);

            // Connect to LibreOffice server
            // URL format explained here: https://wiki.openoffice.org/wiki/Documentation/DevGuide/ProUNO/Starting_OpenOffice.org_in_Listening_Mode
            string connectionString = string.Format("uno:socket,host={0},port={1};urp;StarOffice.ComponentContext", serverHost, serverPort);
            object initialObject = urlResolver.resolve(connectionString);

            // Retrieve remote service manager:
            XComponentContext remoteContext = (XComponentContext)initialObject;
            XMultiComponentFactory remoteComponentFactory = remoteContext.getServiceManager();

            // Request our own instance of LibreOffice Writer from the server:
            object oDesktop = remoteComponentFactory.createInstanceWithContext("com.sun.star.frame.Desktop", remoteContext);
            XComponentLoader xCLoader = (XComponentLoader)oDesktop;

            var loaderArgs = new PropertyValue[1];
            loaderArgs[0] = new PropertyValue();
            loaderArgs[0].Name = "Hidden";  // No GUI
            loaderArgs[0].Value = new Any(true);

            string inputFileUrl = inputFile.AbsoluteUri;

            // WARNING: LibreOffice .NET API uses .NET Remoting and does NOT throw clean and actionable errors,
            //          so, if server crashes or input file is locked or whatever, you will get nulls after casting.
            //          For example, `(XStorable)xCLoader` cast may produce null.
            //          This is counter-intuitive, I know; be more careful in production-grade code and check for nulls after every cast.

            XStorable xStorable = (XStorable)xCLoader.loadComponentFromURL(inputFileUrl, "_blank", 0, loaderArgs);

            var writerArgs = new PropertyValue[2];
            writerArgs[0] = new PropertyValue();
            writerArgs[0].Name = "Overwrite";  // Overwrite outputFile if it already exists
            writerArgs[0].Value = new Any(true);
            writerArgs[1] = new PropertyValue();
            writerArgs[1].Name = "FilterName";
            writerArgs[1].Value = new Any("writer_pdf_Export");  // Important! This filter is doing all PDF stuff.

            string outputFileUrl = outputFile.AbsoluteUri;

            xStorable.storeToURL(outputFileUrl, writerArgs);

            XCloseable xCloseable = (XCloseable)xStorable;
            if (xCloseable != null)
            {
                xCloseable.close(false);
            }
        }
    }

    private static string GetLibreOfficePath()
    {
        using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\LibreOffice\UNO\InstallPath"))
        {
            return key.GetValue(null) as string;
        }
    }
}
