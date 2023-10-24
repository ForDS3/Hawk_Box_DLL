using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using MSOneChart;
using System.Windows.Forms;
using NationalInstruments.TestStand.Interop.API;
using System.Drawing;
using System.Data;
using System.Net.NetworkInformation;
using System.Windows.Forms.VisualStyles;

namespace Hawk
{
    public class HawkBox
    {
        public void GetConfig(SequenceContext seqContext, String CfgPath, out String Project, out String TargetIP, out String FrontCameraConvertTool, out String FrontCameraConvertToolArguments, out String RearCameraConvertTool, out String RearCameraConvertToolArguments, out String IQtool, out String IQtoolArguments, out bool errorOccurred, out int errorCode, out String errorMsg, out String ConvertToolFold, out String FrontCameraConvertImageFold, out String RearCameraConvertImageFold, out String FrontCameraInputImageFold, out String RearCameraInputImageFold, out String CaptureScriptFold, out String AutoMTFToolFold, out String OneChartToolFold, out String RGBSNToolFold, out String IRSNToolFold,
            out float FFCAnchor, out float RFCAnchor)
        {
            errorOccurred = false;
            errorCode = 0;
            errorMsg = String.Empty;
            Project = String.Empty;
            TargetIP = String.Empty;
            FrontCameraConvertTool = String.Empty;
            FrontCameraConvertToolArguments = String.Empty;
            RearCameraConvertTool = String.Empty;
            RearCameraConvertToolArguments = String.Empty;
            IQtool = String.Empty;
            IQtoolArguments = String.Empty;
            FFCAnchor = 0;
            RFCAnchor = 0;
            ConvertToolFold = String.Empty;
            FrontCameraConvertImageFold = String.Empty;
            RearCameraConvertImageFold = String.Empty;
            FrontCameraInputImageFold = String.Empty;
            RearCameraInputImageFold = String.Empty;
            CaptureScriptFold = String.Empty;
            AutoMTFToolFold = String.Empty;
            OneChartToolFold = String.Empty;
            RGBSNToolFold = String.Empty;
            IRSNToolFold = String.Empty;
            try
            {
                string filePath = CfgPath;
                List<string> lines = new List<string>();
                using (StreamReader reader = new StreamReader(filePath + @"\PJTconfig.txt"))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        lines.Add(line.Split('"')[1]);
                    }
                }
                Project = lines[0];
                TargetIP = lines[1];
                FrontCameraConvertTool = lines[2];
                FrontCameraConvertToolArguments = lines[3];
                RearCameraConvertTool = lines[4];
                RearCameraConvertToolArguments = lines[5];
                IQtool = lines[6];
                IQtoolArguments = lines[7];
                FFCAnchor = Convert.ToSingle(lines[8]);
                RFCAnchor = Convert.ToSingle(lines[9]);
                ConvertToolFold = @"C:\CameraDVEF\" + Project + @"\Hawk\MatlabTool";
                FrontCameraConvertImageFold = @"C:\CameraDVEF\" + Project + @"\Hawk\ConvertedImage\FFC";
                RearCameraConvertImageFold = @"C:\CameraDVEF\" + Project + @"\Hawk\ConvertedImage\RFC";
                FrontCameraInputImageFold = @"C:\CameraDVEF\" + Project + @"\Hawk\CapturedImage\FFC";
                RearCameraInputImageFold = @"C:\CameraDVEF\" + Project + @"\Hawk\CapturedImage\RFC";
                CaptureScriptFold = @"C:\CameraDVEF\" + Project + @"\Hawk\CaptureScript";
                AutoMTFToolFold = @"C:\CameraDVEF\" + Project + @"\Hawk\AnalysisTool\AutoMTF";
                OneChartToolFold = @"C:\CameraDVEF\" + Project + @"\Hawk\AnalysisTool\OneChart";
                RGBSNToolFold = @"C:\CameraDVEF\" + Project + @"\Hawk\CaptureScript\DumpTools\RGB";
                IRSNToolFold = @"C:\CameraDVEF\" + Project + @"\Hawk\CaptureScript\DumpTools\IR";
            }
            catch(COMException e)
            {
                errorOccurred = true;
                errorCode = e.ErrorCode;
                errorMsg = e.Message;
            }
        }
        public void NetworkConnectionDetectAndDeviceSNRead(SequenceContext seqContext, String TargetIP, String CaptureImageScriptPath, out String DeviceSN,
out bool errorOccurred, out int errorCode, out String errorMsg)
        {
            errorOccurred = false;
            errorCode = 0;
            errorMsg = String.Empty;
            var host = TargetIP;
            DeviceSN = String.Empty;
            int countdown = 60;
            bool connection = false;
            while (true)
            {
                while (countdown > 0)
                {
                    Console.WriteLine("Time remaining: " + countdown + " seconds");
                    System.Threading.Thread.Sleep(1000); ; // Wait for 1 second
                    countdown--;
                    if (IsPingSuccessful(host))
                    {

                        string filePath = CaptureImageScriptPath + @"\DUTSN.log";

                        if (File.Exists(filePath))
                        {
                            File.Delete(filePath);
                        }

                        ProcessStartInfo startInfo = new ProcessStartInfo();
                        //ProcessStartInfo startInfo = new ProcessStartInfo();
                        startInfo.CreateNoWindow = false;
                        startInfo.UseShellExecute = true;
                        startInfo.FileName = "StartGetDutSN.bat";
                        //startInfo.Arguments = "";
                        startInfo.WindowStyle = ProcessWindowStyle.Normal;
                        startInfo.WorkingDirectory = CaptureImageScriptPath.Replace(@"\\", @"\");

                        Process p = Process.Start(startInfo);
                        do
                        {
                            System.Threading.Thread.Sleep(10);
                        }
                        while (!p.HasExited);
                        string text = File.ReadAllText(filePath);
                        DeviceSN = text.Split('\n')[1].Substring(0, 14);
                        File.Delete(filePath);
                        connection = true;
                        break;
                    }
                    if (countdown == 0)
                    {
                        MessageBox.Show("Check Target IP Address and connection");
                        break;
                    }
                }
                if (connection)
                {
                    break;
                }
                if (countdown == 0)
                {
                    MessageBox.Show("Check Target IP Address and connection");
                    break;
                }
            }
        }
        static bool IsPingSuccessful(string host)
        {
            try
            {
                var ping = new Ping();
                var reply = ping.Send(host);
                return reply.Status == IPStatus.Success;
            }
            catch
            {
                return false;
            }
        }

        public void AutoMTF(SequenceContext seqContext,String IQtoolArguments, String convertedBmpPath, String uuTSN, String AutoMTFTool, out bool errorOccurred, out int errorCode, out String errorMsg)
        {
            errorOccurred = false;
            errorCode = 0;
            errorMsg = String.Empty;

            try
            {
                String bmpImageDirectoryPath = Path.Combine(convertedBmpPath, uuTSN, "RAW_BMPS");
                //File.path
                //RunMatLabtool
                ProcessStartInfo startInfo = new ProcessStartInfo();
                //ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = false;
                startInfo.UseShellExecute = true;
                //startInfo.RedirectStandardOutput = true;
                // 输出错误
                //startInfo.RedirectStandardError = true;
                startInfo.FileName = "_AutoMTF.exe";
                //startInfo.Arguments = "";
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                startInfo.WorkingDirectory = AutoMTFTool.Replace(@"\\", @"\");
                startInfo.Arguments = bmpImageDirectoryPath + " "+ IQtoolArguments;


                Process p = Process.Start(startInfo);
                do
                {
                    System.Threading.Thread.Sleep(10);
                }
                while (!p.HasExited);
            }
            catch (COMException e)
            {
                errorOccurred = true;
                errorMsg = e.Message;
                errorCode = e.ErrorCode;
            }
        }

        public void ParseAutoMTF(SequenceContext seqContext, String ConvertedBmpPath, String UUTSN_WithTimeStamp, out double Centermtf, out double Cornermtf, out bool errorOccurred, out int errorCode, out String errorMsg)
        {

            errorOccurred = false;
            errorCode = 0;
            errorMsg = String.Empty;
            Centermtf = 0;
            Cornermtf = 0;

            try
            {
                String bmpImageDirectoryPath = Path.Combine(ConvertedBmpPath, UUTSN_WithTimeStamp, "RAW_BMPS");
                String[] csvFiles = Directory.GetFiles(bmpImageDirectoryPath, "*.csv");

                DataTable dtIR_Center_PointA = new DataTable();

                System.IO.FileStream fs = new System.IO.FileStream(csvFiles[0], System.IO.FileMode.Open);
                System.IO.StreamReader sr = new System.IO.StreamReader(fs, Encoding.GetEncoding("gb2312"));
                string temptext = "";
                bool isFirst = true;
                while ((temptext = sr.ReadLine()) != null)
                {
                    string[] arr = temptext.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                    if (isFirst)
                    {
                        foreach (string s in arr)
                        {
                            dtIR_Center_PointA.Columns.Add(s);
                        }
                        isFirst = false;
                    }
                    else
                    {
                        DataRow dr = dtIR_Center_PointA.NewRow();
                        for (int i = 0; i < dtIR_Center_PointA.Columns.Count; i++)
                        {
                            dr[i] = i < arr.Length ? arr[i] : "";
                        }
                        dtIR_Center_PointA.Rows.Add(dr);
                    }
                }

                Centermtf = Convert.ToDouble(dtIR_Center_PointA.Rows[28][1]);
                Cornermtf = Convert.ToDouble(dtIR_Center_PointA.Rows[29][1]);
            }

            catch (COMException e)
            {
                errorOccurred = true;
                errorMsg = e.Message;
                errorCode = e.ErrorCode;
            }
        }

        public void OneChartTest(SequenceContext seqContext, String convertedBmpPath, String uuTSN, String ReportPath, float FFCAncor ,out bool errorOccurred, out int errorCode, out String errorMsg, out float Rotation, out float Pan, out float Tilt, out float OneChartMTF)
        {
            String[] CheckcsvFiles = Directory.GetFiles(ReportPath, "data.csv", SearchOption.AllDirectories);
            if(CheckcsvFiles.Length>0)
            {
                foreach (string f in CheckcsvFiles)
                {
                    File.Delete(f);
                }
            }
            errorOccurred = false;
            errorCode = 0;
            errorMsg = String.Empty;
            Rotation = 0;
            Pan = 0;
            Tilt = 0;
            OneChartMTF = 0;
            try
            {
                Rectangle Crop = new Rectangle(0, 0, 0, 0);
                AllDataReport onecharts = new MSOneChart.AllDataReport();
                MTFDataReport mtfreport = new MSOneChart.MTFDataReport();
                onecharts.AnchorThresh = FFCAncor;
                onecharts.setMTFLocation(AllDataReport.MTFTarget.LARGE, 10);
                onecharts.setMTFLocation(AllDataReport.MTFTarget.SMALL, 10);
                onecharts.setGeoSetup(500, 37);
                onecharts.GeoAligned = true;
                onecharts.SaveIndReport = false;
                onecharts.ContrastMTF = false;
                onecharts.Channel = 1;
                onecharts.setPath(ReportPath);
                onecharts.setSESFRFreq(AllDataReport.MTFTarget.LARGE, MTFDataReport.SLANTEDGESFRNY, 4);
                onecharts.setSESFRFreq(AllDataReport.MTFTarget.SMALL, MTFDataReport.SLANTEDGESFRNY, 4);
                String bmpImageDirectoryPath = Path.Combine(convertedBmpPath, uuTSN, "RAW_BMPS");
                String[] bmpFiles = Directory.GetFiles(bmpImageDirectoryPath, "*.bmp");
                onecharts.openOCImage(bmpFiles[0], Crop, System.Drawing.RotateFlipType.RotateNoneFlipNone, out OneChartImage ocImg);
                onecharts.processOCImage(ocImg, false, out bool TestResult);
                String[] csvFiles = Directory.GetFiles(ReportPath, "data.csv", SearchOption.AllDirectories);
                System.IO.FileStream fs = new System.IO.FileStream(csvFiles[0], System.IO.FileMode.Open);
                System.IO.StreamReader sR = new System.IO.StreamReader(fs, Encoding.GetEncoding("gb2312"));
                List<string> result = new List<string>();
                string nextLine;
                while ((nextLine = sR.ReadLine()) != null)
                {
                    foreach (string i in nextLine.Split(','))
                    {
                        result.Add(i);
                    }

                }
                sR.Close();
                OneChartMTF = Convert.ToSingle(result[304]);
                Rotation = Convert.ToSingle(result[460]);
                Pan = Convert.ToSingle(result[461]);
                Tilt = Convert.ToSingle(result[462]);

                foreach (string f in csvFiles)
                {
                    File.Delete(f);
                }
            }
            catch(COMException e)
            {
                errorOccurred = true;
                errorCode = e.ErrorCode;
                errorMsg = e.Message;
            }


        }
        public void CallFFCCaptureBatchFile(SequenceContext seqContext, String CaptureImageScriptPath, String CapturedRawPath, out bool errorOccurred, out int errorCode, out String errorMsg)
        {
            errorOccurred = false;
            errorCode = 0;
            errorMsg = String.Empty;


            try
            {
                //Capture Script
                ProcessStartInfo startInfo = new ProcessStartInfo();
                //ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = false;
                startInfo.UseShellExecute = true;
                startInfo.FileName = "StartFFCCapture.bat";
                //startInfo.Arguments = "";
                startInfo.WindowStyle = ProcessWindowStyle.Normal;
                startInfo.WorkingDirectory = CaptureImageScriptPath.Replace(@"\\", @"\");
                //startInfo.Arguments = convertedBmpPath.Replace(@"\\", @"\") + "\\" + uutSN_WithTimeStamp + " 644 604 raw 1 1";


                Process p = Process.Start(startInfo);
                do
                {
                    System.Threading.Thread.Sleep(10);
                }
                while (!p.HasExited);

                if (!(Directory.GetFiles(CapturedRawPath).Length > 0))
                {
                    MessageBox.Show("Capture Image failed, no image been generaged");
                }
            }
            catch (COMException e)
            {
                errorOccurred = true;
                errorMsg = e.Message;
                errorCode = e.ErrorCode;
            }
        }
        public void CallRFCCaptureBatchFile(SequenceContext seqContext, String CaptureImageScriptPath, String CapturedRawPath, out bool errorOccurred, out int errorCode, out String errorMsg)
        {
            errorOccurred = false;
            errorCode = 0;
            errorMsg = String.Empty;


            try
            {
                //Capture Script
                ProcessStartInfo startInfo = new ProcessStartInfo();
                //ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = false;
                startInfo.UseShellExecute = true;
                startInfo.FileName = "StartRFCCapture.bat";
                //startInfo.Arguments = "";
                startInfo.WindowStyle = ProcessWindowStyle.Normal;
                startInfo.WorkingDirectory = CaptureImageScriptPath.Replace(@"\\", @"\");
                //startInfo.Arguments = convertedBmpPath.Replace(@"\\", @"\") + "\\" + uutSN_WithTimeStamp + " 644 604 raw 1 1";


                Process p = Process.Start(startInfo);
                do
                {
                    System.Threading.Thread.Sleep(10);
                }
                while (!p.HasExited);

                if (!(Directory.GetFiles(CapturedRawPath).Length > 0))
                {
                    MessageBox.Show("Capture Image failed, no image been generaged");
                }
            }
            catch (COMException e)
            {
                errorOccurred = true;
                errorMsg = e.Message;
                errorCode = e.ErrorCode;
            }
        }

        public void ConvertRawToBmpForFFC(SequenceContext seqContext, String scanedSN, String FrontCameraConvertTool, String FrontCameraConvertToolArguments, String sourceRawPath, String convertedBmpPath, String ToolBatPath, out String reportText, out bool errorOccurred, out int errorCode, out String errorMsg, out String uutSN_WithTimeStamp)
        {
            reportText = String.Empty;
            errorOccurred = false;
            errorCode = 0;
            errorMsg = String.Empty;
            uutSN_WithTimeStamp = String.Empty;
            //SN_TimeStamp folder
            String DestinationDirectory = String.Empty;

            try
            {
                // INSERT YOUR SPECIFIC TEST CODE HERE

                // The following code shows how to access properties and variables via the TestStand API
                // PropertyObject propertyObject = seqContext.AsPropertyObject();
                // String lastUserName = propertyObject.GetValString("StationGlobals.TS.LastUserName", 0);
                //reportText = "this is Rich's string from code module";
                uutSN_WithTimeStamp = scanedSN + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                DestinationDirectory = Path.Combine(convertedBmpPath, uutSN_WithTimeStamp);
                //create new folder
                if (!Directory.Exists(DestinationDirectory))
                {
                    Directory.CreateDirectory(DestinationDirectory);
                }

                //move all raw file to destination folder, typical only 1 .raw file
                DirectoryInfo sourceFolderInfo = new DirectoryInfo(sourceRawPath);

                String tempFileName = String.Empty;//for store the raw file name

                foreach (FileInfo NextFile in sourceFolderInfo.GetFiles())
                {
                    File.Move(NextFile.FullName, Path.Combine(DestinationDirectory, NextFile.Name));
                    tempFileName = Path.Combine(DestinationDirectory, NextFile.Name);
                }
                //File.path

                //RunMatLabtool
                ProcessStartInfo startInfo = new ProcessStartInfo();
                //ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = false;
                startInfo.UseShellExecute = true;
                //startInfo.RedirectStandardOutput = true;
                // 输出错误
                //startInfo.RedirectStandardError = true;
                startInfo.FileName = FrontCameraConvertTool;
                //startInfo.Arguments = "";
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                startInfo.WorkingDirectory = ToolBatPath.Replace(@"\\", @"\");
                startInfo.Arguments = convertedBmpPath.Replace(@"\\", @"\") + "\\" + uutSN_WithTimeStamp + " "+ FrontCameraConvertToolArguments;


                Process p = Process.Start(startInfo);
                do
                {
                    System.Threading.Thread.Sleep(10);
                }
                while (!p.HasExited);
            }

            catch (COMException e)
            {
                errorOccurred = true;
                errorMsg = e.Message;
                errorCode = e.ErrorCode;
            }
        }

        public void ConvertRawToBmpForRFC(SequenceContext seqContext, String scanedSN, String sourceRawPath, String RearCameraConvertTool, String RearCameraConvertToolArguments, String convertedBmpPath, String ToolBatPath, out String reportText, out bool errorOccurred, out int errorCode, out String errorMsg, out String uutSN_WithTimeStamp)
        {
            reportText = String.Empty;
            errorOccurred = false;
            errorCode = 0;
            errorMsg = String.Empty;
            uutSN_WithTimeStamp = String.Empty;
            //SN_TimeStamp folder
            String DestinationDirectory = String.Empty;

            try
            {
                uutSN_WithTimeStamp = scanedSN + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                DestinationDirectory = Path.Combine(convertedBmpPath, uutSN_WithTimeStamp);
                //create new folder
                if (!Directory.Exists(DestinationDirectory))
                {
                    Directory.CreateDirectory(DestinationDirectory);
                }
                DirectoryInfo sourceFolderInfo = new DirectoryInfo(sourceRawPath);

                String tempFileName = String.Empty;//for store the raw file name

                foreach (FileInfo NextFile in sourceFolderInfo.GetFiles())
                {
                    File.Move(NextFile.FullName, Path.Combine(DestinationDirectory, NextFile.Name));
                    tempFileName = Path.Combine(DestinationDirectory, NextFile.Name);
                }
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = false;
                startInfo.UseShellExecute = true;
                startInfo.FileName = RearCameraConvertTool;
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                startInfo.WorkingDirectory = ToolBatPath.Replace(@"\\", @"\");
                startInfo.Arguments = convertedBmpPath.Replace(@"\\", @"\") + "\\" + uutSN_WithTimeStamp + " "+ RearCameraConvertToolArguments;
                Process p = Process.Start(startInfo);
                do
                {
                    System.Threading.Thread.Sleep(10);
                }
                while (!p.HasExited);
            }

            catch (COMException e)
            {
                errorOccurred = true;
                errorMsg = e.Message;
                errorCode = e.ErrorCode;
            }
        }

        public void ShowImage(String convertedBmpPath, String uuTSN, out bool errorOccurred, out int errorCode, out String errorMsg)
        {
            errorOccurred = false;
            errorCode = 0;
            errorMsg = String.Empty;

            try
            {
                String bmpImageDirectoryPath = Path.Combine(convertedBmpPath, uuTSN, "RAW_BMPS");
                String[] bmpFiles = Directory.GetFiles(bmpImageDirectoryPath, "*.bmp");
                ProcessStartInfo startInfo = new ProcessStartInfo("cmd.exe");
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                startInfo.Arguments = @"cmd /c %systemroot%\system32\rundll32.exe " + "\"" + @"%programfiles%\windows photo viewer\PhotoViewer.dll"
                                      + "\"" + @",imageview_fullscreen " + bmpFiles[0];
                Process p = Process.Start(startInfo);

            }
            catch (COMException e)
            {
                errorOccurred = true;
                errorMsg = e.Message;
                errorCode = e.ErrorCode;
            }
        }

        public void WriteLogs(string convertedBmpPath, String uuTSN, float CenterMTF, float CornerMTF,float OneChartMTF, float Rotation, float Pan, float Tilt, out bool errorOccurred, out int errorCode, out String errorMsg)
        {
            errorOccurred = false;
            errorCode = 0;
            errorMsg = String.Empty;
            List<string> result = new List<string>();
            result.Add("CenterMTF:"+Convert.ToString(CenterMTF));
            result.Add("CornerMTF:" + Convert.ToString(CornerMTF));
            result.Add("OneChartMTF:" + Convert.ToString(OneChartMTF));
            result.Add("Rotation:" + Convert.ToString(Rotation));
            result.Add("Pan:" + Convert.ToString(Pan));
            result.Add("Tilt:" + Convert.ToString(Tilt));

            string Filename = convertedBmpPath + @"\"+uuTSN + @"\result.log";

            if (!File.Exists(Filename))
            {
                FileStream fileStream = new FileStream(Filename, FileMode.Create, FileAccess.Write);
                fileStream.Close();
                StreamWriter sw = new StreamWriter(Filename);
                foreach (string i in result)
                {
                    sw.Write(i+"\n");
                }
                sw.Close();
            }
        }
    }
}
