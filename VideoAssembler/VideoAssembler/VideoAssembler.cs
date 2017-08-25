using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using QTOLibrary;
using QTOControlLib;


namespace VideoAssemblerCMD
{
	class Program
	{
		
		///<summary>
		///This is a command line application that combines video files and exports it.
		///This application takes two arguments - source folderpath (or text file with a file paths) and destination path.
		///</summary>
		///<param name="args"></param>
		static void Main(string[] args)
		{
			
			//initiate a timer to check the elapsed time
			Stopwatch timer = new Stopwatch();
			timer.Start();
			
			
			//retry flag
			bool retry = true;
			
			while(retry)
			{
				//create a type of quick time
				Type QtType = Type.GetTypeFromProgID("QuickTimePlayerLib.QuickTimePlayerApp");
				//use Activator to create an instance of quicktime
				dynamic qtPlayerApp = Activator.CreateInstance(QtType);
				
				
				try
				{
					string sourcePath = args[0];
					string destPath = args[1];
					
					Thread.Sleep(5000);
					
					//Create two quick timeplayers from the app
					var qtPlayerSrc = qtPlayerApp.Players[1];
					qtPlayerApp.Players.Add();
					var qtPlayerDest = qtPlayerApp.Players(qtPlayerApp.Players.Count);
					
					
					//Create a new empty movie in dest player
					qtPlayerDest.QTControl.CreateNewMovie(true);
					
					//empty list of strings to hold the movie urls
					List<string> movies = new List<string>();
					
					//check for sourcePath as txt file or folder
					if(sourcePath.EndsWith(".txt"))
					{
						using(StreamReader reader = new StreamReader(sourcePath))
						{
							string line;
							
							while((line = reader.ReadLine()) != null)
							{
								movies.Add(line);
							}
						}
						
					}
					else
					{
						//get all the .mov files from the souce folder
						string[] mov_list = Directory.GetFiles(sourcePath, "*.mov");
						
						
						foreach(string file in mov_list)
						{
							movies.Add(file);
						}
					}
					
					int count = 0;
					
					//iterate over the files
					foreach(string mov in movies)
					{
						
						count++;
						Console.WriteLine("Copying "+ mov);
						
						char[] charToTrim = {'"', '\''};
						
						if (mov.StartsWith("\"") || mov.StartsWith("'"))
						{
							qtPlayerSrc.QTControl.URL = mov.Trim(charToTrim);
						}						
						else
						{
							qtPlayerSrc.QTControl.URL = mov;
						}
						
						//select the whole file to copy
						if (qtPlayerSrc.QTControl.Movie != null)
						{
							qtPlayerSrc.QTControl.Movie.SelectAll();
						}
						
						//Use InsertSegment() to copy the movie over to qtPlayer1
						qtPlayerDest.QTControl.Movie.InsertSegment(qtPlayerSrc.QTControl.Movie,
						qtPlayerSrc.QTControl.Movie.SelectionStart,
						qtPlayerSrc.QTControl.Movie.SelectionEnd,
						qtPlayerDest.QTControl.Movie.Duration);
						
					}
					
					Console.WriteLine("Combined " + count + " video files");
					
					//Rewind the movie in qtPlayer1 and set the SelectionDuraion to 0
					qtPlayerDest.QTControl.Movie.Rewind();
					qtPlayerDest.QTControl.Movie.SelectionDuraion = 0;
					
					
					//Export the destination movie and close the player
					QTQuickTime qt = new QTQuickTime();
					qt = qtPlayerDest.QTControl.QuickTime;
					qt.Exporters.Add();
					
					QTExporter exp = qt.Exporters[1];
					
					//Define the settings for the exporter - source, destination, type
					exp.SetDataSource(qtPlayerDest.QTControl.Movie);
					exp.TypeName = "QuickTime Movie";
					exp.DestinationFileName = destPath;
					
					
					//Export begins
					Console.WriteLine("Exporting...");
					exp.BeginExport();
					
					qtPlayerDest.Close();
					
					//Quit QuickTime
					qtPlayerApp.Quit();
					Console.WriteLine("Exported to destination: " + destPath);
					
					//if execution is successful, set retry flag to false.
					retry = false;
					
				}
				catch(COMException ce)
				{
					string msg;
					QTUtils qtu = new QTUtils();
					
					msg = "Error Occurred !";
					msg += "\nError Code: " + ce.ErrorCode;
					msg += "\nQT Error Code: " + qtu.QTErrorFromErrorCode(ce.ErrorCode);
					msg += ce;
					Console.WriteLine(msg);
					
					Console.WriteLine("\nRestarting...");
					//if execution fails, set retry flag to true
					retry = true;
				}
				
				
			}
			
			Console.WriteLine("Execution Time: " + timer.Elapsed);
			
			
		}
	}
	
}