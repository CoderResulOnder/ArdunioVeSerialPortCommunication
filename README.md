
        #region Arduino ve Seri Port Bağlantı İşlemleri

        #region Ardunio ve RAMAN portlarını tespit ile ilgili classlar
        internal class ProcessConnection
        {

            public static ConnectionOptions ProcessConnectionOptions()
            {
                ConnectionOptions options = new ConnectionOptions();
                options.Impersonation = ImpersonationLevel.Impersonate;
                options.Authentication = AuthenticationLevel.Default;
                options.EnablePrivileges = true;
                return options;
            }

            public static ManagementScope ConnectionScope(string machineName, ConnectionOptions options, string path)
            {
                ManagementScope connectScope = new ManagementScope();
                connectScope.Path = new ManagementPath(@"\\" + machineName + path);
                connectScope.Options = options;
                connectScope.Connect();
                return connectScope;
            }
        }

        public class COMPortInfo
        {
            public string Name { get; set; }
            public string Description { get; set; }

            public COMPortInfo() { }

            public static List<COMPortInfo> GetCOMPortsInfo()
            {
                List<COMPortInfo> comPortInfoList = new List<COMPortInfo>();

                ConnectionOptions options = ProcessConnection.ProcessConnectionOptions();
                ManagementScope connectionScope = ProcessConnection.ConnectionScope(Environment.MachineName, options, @"\root\CIMV2");

                ObjectQuery objectQuery = new ObjectQuery("SELECT * FROM Win32_PnPEntity WHERE ConfigManagerErrorCode = 0");
                ManagementObjectSearcher comPortSearcher = new ManagementObjectSearcher(connectionScope, objectQuery);

                using (comPortSearcher)

                {
                    string caption = null;
                    foreach (ManagementObject obj in comPortSearcher.Get())
                    {
                        if (obj != null)
                        {
                            object captionObj = obj["Caption"];
                            if (captionObj != null)
                            {
                                caption = captionObj.ToString();
                                if (caption.Contains("(COM"))
                                {
                                    COMPortInfo comPortInfo = new COMPortInfo();
                                    comPortInfo.Name = caption.Substring(caption.LastIndexOf("(COM")).Replace("(", string.Empty).Replace(")",
                                                                         string.Empty);
                                    comPortInfo.Description = caption;
                                    comPortInfoList.Add(comPortInfo);
                                }
                            }
                        }
                    }
                }
                return comPortInfoList;
            }
        }
        #endregion



        public void portScan() //Bağlı olan portları listeler
        {
           
            foreach (COMPortInfo comPort in COMPortInfo.GetCOMPortsInfo())
            {
                string des = ((comPort.Description).ToString()).ToLower();
                string portName = (comPort.Name).ToString();
              

                if (des.Contains("usb seri cihaz") || des.Contains("usb serial device"))//usb seri cihaz
                {
                    try
                    {
                       
                        if (!serialPortArduino.IsOpen && ArdunioKontrol)
                        {
                            serialPortArduino.PortName = portName;
                            serialPortArduino.Close();
                            serialPortArduino.Open();
                            serialPortArduino.BaudRate = 9600;
                            serialPortArduino.WriteLine("2");
                            OlcumeHazir();
                        }
                    }
                    catch (Exception ex)
                    {
                       
                    }
                   
                }
                else if (des.Contains("usb serial port"))
                {//Buradan diger spectro portu ögrenilir
                    string a = Regex.Replace(portName, "[^0-9]+", string.Empty);

                    dwComPort = Convert.ToUInt32(a);
                    dwBaudrate = 9600;
                       
                }

            }
        
        }

     

        public void closePort() //Portu kapat.
        {
            if (serialPortArduino.IsOpen && ArdunioKontrol) 
            {
                serialPortArduino.WriteLine("0");
                serialPortArduino.Close();
            }
        }

        #endregion
        
