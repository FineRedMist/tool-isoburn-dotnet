using System;
using System.Collections.Generic;
using System.Text;
using IMAPI2;
using System.Runtime.InteropServices;
using System.Threading;

namespace ISOBurn
{
    class Program
    {
        #region Error codes
        const int ERR_INVALID_ARGS                  = 1;
        const int ERR_FILE_NOT_FOUND                = 2;
        const int ERR_DRIVE_NOT_SUPPORTED           = 3;
        const int ERR_DRIVE_IN_USE                  = 4;
        const int ERR_DRIVE_FAILED_ACQUIRE          = 5;
        const int ERR_DISK_NOT_BLANK                = 6;
        const int ERR_DRIVE_RECORDER_NOT_SUPPORTED  = 7;
        const int ERR_DRIVE_DISK_NOT_SUPPORTED      = 8;
        const int ERR_DRIVE_NOT_FOUND               = 9;
        #endregion

        /// <summary>
        /// Creates an IStream from the given file
        /// </summary>
        /// <param name="pszFile">File to create a stream from</param>
        /// <param name="grfMode">Flags on how to open the file</param>
        /// <param name="ppstm">IStream object returned</param>
        /// <returns>An HRESULT indicating pass or fail</returns>
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
        static extern int SHCreateStreamOnFile(string pszFile, uint grfMode, out IMAPI2.IStream ppstm);

        /// <summary>
        /// Wrapper for logging in case I need to make sure it responds faster than a write to a console window
        /// </summary>
        /// <param name="LogLine">Format string to log</param>
        /// <param name="list">Parameters to the format string</param>
        public static void Log(string LogLine, params object [] list)
        {
            Console.WriteLine(LogLine, list);
        }

        [STAThread]
        static int Main(string[] args)
        {
            int Result = 0; 
            bool HasExclusive = false;

            // First and second arguments
            string Drive = null, ISOImage = null;

            MsftDiscMaster2Class DiscMaster = new MsftDiscMaster2Class();
            MsftDiscRecorder2Class DiscRecorder = null;
            DDiscFormat2DataEvents_UpdateEventHandler EventHandler = new DDiscFormat2DataEvents_UpdateEventHandler(DiscFormatData_Update);

            // Check arguments.
            if (args.Length < 2)
            {
                Log("Usage: ISOBurn <DrivePath> <ISOImage>");
                Log("\tISOBurn e: d:\\summerpics.iso");
                Result = ERR_INVALID_ARGS;
                goto Exit;
            }

            ISOImage = args[1];

            // If the file doesn't exist, waste no more time
            if (!System.IO.File.Exists(args[1]))
            {
                Console.WriteLine("File not found: {0}", ISOImage);
                Result = ERR_FILE_NOT_FOUND;
                goto Exit;
            }
            
            // Find the corresponding disc recorder for the drive
            for (int i = 0; i < DiscMaster.Count; ++i)
            {
                MsftDiscRecorder2Class DR = new MsftDiscRecorder2Class();
                DR.InitializeDiscRecorder(DiscMaster[i]);
                foreach (string s in DR.VolumePathNames)
                {
                    // Indicate the drive letter the system is using and set the disc recorder
                    if (s.StartsWith(args[0], StringComparison.CurrentCultureIgnoreCase))
                    {
                        Log("Using drive {0} for burning from ID {1}", s, DiscMaster[i]);
                        Drive = s;
                        DiscRecorder = DR;
                        break;
                    }
                }
            }

            // If we didn't set DiscRecorder, then no drive was found
            if (DiscRecorder == null)
            {
                Log("The drive {0} to burn to was not found.", Drive);
                Result = ERR_DRIVE_NOT_FOUND;
                goto Exit;
            }

            // Ensure the drive supports writing.
            List<IMAPI_FEATURE_PAGE_TYPE> Features = new List<IMAPI_FEATURE_PAGE_TYPE>();
            foreach (IMAPI_FEATURE_PAGE_TYPE IFPT in DiscRecorder.SupportedFeaturePages)
            {
                Features.Add(IFPT);
            }

            // Is there a better way to do this?  Doesn't look like a bitfield based on the enums (saddened)
            if (!Features.Contains(IMAPI_FEATURE_PAGE_TYPE.IMAPI_FEATURE_PAGE_TYPE_BD_WRITE) &&
                !Features.Contains(IMAPI_FEATURE_PAGE_TYPE.IMAPI_FEATURE_PAGE_TYPE_CD_RW_MEDIA_WRITE_SUPPORT) &&
                !Features.Contains(IMAPI_FEATURE_PAGE_TYPE.IMAPI_FEATURE_PAGE_TYPE_CDRW_CAV_WRITE) &&
                !Features.Contains(IMAPI_FEATURE_PAGE_TYPE.IMAPI_FEATURE_PAGE_TYPE_DOUBLE_DENSITY_CD_R_WRITE) &&
                !Features.Contains(IMAPI_FEATURE_PAGE_TYPE.IMAPI_FEATURE_PAGE_TYPE_DOUBLE_DENSITY_CD_RW_WRITE) &&
                !Features.Contains(IMAPI_FEATURE_PAGE_TYPE.IMAPI_FEATURE_PAGE_TYPE_DVD_DASH_WRITE) &&
                !Features.Contains(IMAPI_FEATURE_PAGE_TYPE.IMAPI_FEATURE_PAGE_TYPE_DVD_PLUS_R) &&
                !Features.Contains(IMAPI_FEATURE_PAGE_TYPE.IMAPI_FEATURE_PAGE_TYPE_DVD_PLUS_R_DUAL_LAYER) &&
                !Features.Contains(IMAPI_FEATURE_PAGE_TYPE.IMAPI_FEATURE_PAGE_TYPE_DVD_PLUS_RW) &&
                !Features.Contains(IMAPI_FEATURE_PAGE_TYPE.IMAPI_FEATURE_PAGE_TYPE_HD_DVD_WRITE) &&
                !Features.Contains(IMAPI_FEATURE_PAGE_TYPE.IMAPI_FEATURE_PAGE_TYPE_CD_MASTERING))
            {
                Log("The drive {0} does not appear to support writing.", Drive);
                Result = ERR_DRIVE_NOT_SUPPORTED;
                goto Exit;
            }

            // Check to see if another application is already using the media and if so bail
            MsftDiscFormat2DataClass DiscFormatData = null;
            if (DiscRecorder.ExclusiveAccessOwner != null && DiscRecorder.ExclusiveAccessOwner.Length != 0)
            {
                Log("The drive {0} is in use by the application: {1}", Drive, DiscRecorder.ExclusiveAccessOwner);
                Result = ERR_DRIVE_IN_USE;
                goto Exit;
            }
            // Try to get exclusive access.  This is important.  In Vista 5728, Media Player causes this to return E_FAIL
            // even if it is just playing music in the background.
            try
            {
                DiscRecorder.AcquireExclusiveAccess(false, "ISOBurn Commandline Tool");
            }
            catch (System.Exception e)
            {
                Log("Failed to acquire exclusive access to the burner: Message: {0}\nStack Trace: {1}", e.Message, e.StackTrace);
                Result = ERR_DRIVE_FAILED_ACQUIRE;
                goto Exit;
            }

            // Disable media change notifications so we don't have anyone else being notified that I'm doing stuff on the drive
            DiscRecorder.DisableMcn();

            // Indicate we have exclusive access.
            HasExclusive = true;

            // Get the disk format which will hopefully let us know if we can write to the disk safely.
            Result = GetDiskFormatData(DiscRecorder, Drive, out DiscFormatData);
            if (Result != 0)
            {
                goto Exit;
            }

            // I would like to get the amount of free space on the media and compare that against the file size, but I'm not sure
            // if the file size of the ISO represents the raw sectors that would be written or whether it might be more tightly 
            // packed.  Also there might be additional sectors written on the disk to identify the image, etc that isn't 
            // represented in the ISO.  These are details I don't know to make this a little bit more robust.

            // Get the image to write
            IMAPI2.IStream Stream = null;
            int Res = SHCreateStreamOnFile(ISOImage, 0x20, out Stream);
            if (Res < 0)
            {
                Log("Opening the source ISO image {0} failed with error: {1}", ISOImage, Res);
                Result = Res;
                goto Exit;
            }

            // Set the client name
            DiscFormatData.ClientName = "ISOBurn Commandline Tool";
            // Add the event handler *WHICH CURRENTLY DOES GENERATE EVENTS*
            DiscFormatData.Update += EventHandler;

            Log("Disk write speed is {0} sectors/second", DiscFormatData.CurrentWriteSpeed);

            // Write the stream
            try
            {
                DiscFormatData.Write(Stream);
                Log("Burn complete of {0} to {1}!", ISOImage, Drive);
            }
            catch (System.Exception e)
            {
                Log("Burn Failed: {0}\n{1}", e.Message, e.StackTrace.ToString());
                goto Exit;
            }

        Exit:
            // Cleanup
            if (DiscRecorder != null && HasExclusive)
            {
                DiscRecorder.EnableMcn();
                DiscRecorder.ReleaseExclusiveAccess();
                DiscRecorder.EjectMedia();
            }

            return Result;
        }

        /// <summary>
        /// Outputs an updated status of the write.  Currently not being called--don't know why.
        /// Documentation at: http://windowssdk.msdn.microsoft.com/en-us/library/ms689023.aspx
        /// </summary>
        /// <param name="obj">An IDiscFormat2Data interface</param>
        /// <param name="prog">An IDiscFormat2DataEventArgs interface</param>
        static void DiscFormatData_Update(object obj, object prog)
        {
            // Update the status progress of the write.
            try
            {
                IDiscFormat2DataEventArgs progress = (IDiscFormat2DataEventArgs)prog;

                string strTimeStatus = "Time: " + progress.ElapsedTime + " / " + progress.TotalTime;

                switch (progress.CurrentAction)
                {
                    case IMAPI2.IMAPI_FORMAT2_DATA_WRITE_ACTION.IMAPI_FORMAT2_DATA_WRITE_ACTION_VALIDATING_MEDIA:
                        strTimeStatus = "Validating media " + strTimeStatus;
                        break;

                    case IMAPI2.IMAPI_FORMAT2_DATA_WRITE_ACTION.IMAPI_FORMAT2_DATA_WRITE_ACTION_FORMATTING_MEDIA:
                        strTimeStatus = "Formatting media " + strTimeStatus;
                        break;

                    case IMAPI2.IMAPI_FORMAT2_DATA_WRITE_ACTION.IMAPI_FORMAT2_DATA_WRITE_ACTION_INITIALIZING_HARDWARE:
                        strTimeStatus = "Initializing Hardware " + strTimeStatus;
                        break;

                    case IMAPI2.IMAPI_FORMAT2_DATA_WRITE_ACTION.IMAPI_FORMAT2_DATA_WRITE_ACTION_CALIBRATING_POWER:
                        strTimeStatus = "Calibrating Power (OPC) " + strTimeStatus;
                        break;

                    case IMAPI2.IMAPI_FORMAT2_DATA_WRITE_ACTION.IMAPI_FORMAT2_DATA_WRITE_ACTION_WRITING_DATA:
                        long totalSectors, writtenSectors;
                        double percentDone;
                        totalSectors = progress.SectorCount;
                        writtenSectors = progress.LastWrittenLba - progress.StartLba;
                        percentDone = writtenSectors * 100;
                        percentDone /= totalSectors;
                        strTimeStatus = "Progress:  " + percentDone.ToString("0.00") + "%  " + strTimeStatus;
                        break;
                    case IMAPI2.IMAPI_FORMAT2_DATA_WRITE_ACTION.IMAPI_FORMAT2_DATA_WRITE_ACTION_FINALIZATION:
                        strTimeStatus = "Finishing the writing " + strTimeStatus;
                        break;
                    case IMAPI2.IMAPI_FORMAT2_DATA_WRITE_ACTION.IMAPI_FORMAT2_DATA_WRITE_ACTION_COMPLETED:
                        strTimeStatus = "Completed the burn.";
                        break;
                    default:
                        strTimeStatus = "Unknown action: " + progress.CurrentAction;
                        break;
                };
                Program.Log(strTimeStatus);
            }
            catch (System.Exception e)
            {
                Program.Log("Update Exception: Message: {0}\nStack Trace: {1}", e.Message, e.StackTrace.ToString());
            }
        }

        /// <summary>
        /// Determines if the format of the media is ok to write and a few other details.  Does some informational output too.
        /// </summary>
        /// <param name="DiscRecorder">DiscRecorder being written with</param>
        /// <param name="Drive">Drive for the DiscRecorder</param>
        /// <param name="DiscFormatData">The MsftDiscFormat2DataClass that will be returned if it is ok to write</param>
        /// <returns>Zero if successful, non-zero otherwise.</returns>
        static int GetDiskFormatData(MsftDiscRecorder2Class DiscRecorder, string Drive, out MsftDiscFormat2DataClass DiscFormatData)
        {
            MsftDiscFormat2DataClass DFD = new MsftDiscFormat2DataClass();
            DFD.Recorder = DiscRecorder;

            DiscFormatData = null;
            if (!DFD.IsRecorderSupported(DiscRecorder))
            {
                Log("The recorder for drive {0} is not supported.", Drive);
                return ERR_DRIVE_RECORDER_NOT_SUPPORTED;
            }

            if (!DFD.IsCurrentMediaSupported(DiscRecorder))
            {
                Log("The media for drive {0} is not supported.", Drive);
                return ERR_DRIVE_DISK_NOT_SUPPORTED;
            }

            // I'm not sure exactly what the difference is between heuristically blank and blank is unless it means that there is room
            // left to be written to and it isn't finalised?
            if (!DFD.MediaHeuristicallyBlank)
            {
                Log("The disk in drive {0} isn't \"heuristically\" blank.  No media was written.", Drive);
                return ERR_DISK_NOT_BLANK;
            }

            // Get the media status properties for the media to for diagnostic purposes
            StringBuilder sb = new StringBuilder();
            bool FirstWritten = false;
            sb.Append("Media status for drive ").Append(Drive).Append(" is ");
            
            for (int i = 1; i <= (int) DFD.CurrentMediaStatus; i <<= 1)
            {
                if ((((int) DFD.CurrentMediaStatus) & i) == i)
                {
                    if (FirstWritten)
                    {
                        sb.Append(" | ");
                    }
                    FirstWritten = true;
                    sb.Append(((IMAPI_FORMAT2_DATA_MEDIA_STATE)i).ToString());
                }
            }
            Log("{0}", sb.ToString());

            // And log the media type for diagnostics
            Log("Media type for drive {0} is {1}", Drive, DFD.CurrentPhysicalMediaType.ToString());

            // Return the result since we are good if we got here.
            DiscFormatData = DFD;
            return 0;
        }
    }
}
