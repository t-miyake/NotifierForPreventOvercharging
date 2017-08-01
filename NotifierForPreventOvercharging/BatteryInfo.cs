using System;
using System.IO;
using System.Runtime.InteropServices;

namespace NotifierForPreventOvercharging
{
    public static class BatteryInfo
    {
        public static BatteryInformation GetBatteryInformation()
        {
            var deviceDataPointer = IntPtr.Zero;
            var queryInfoPointer = IntPtr.Zero;
            var batteryInfoPointer = IntPtr.Zero;
            var batteryWaitStatusPointer = IntPtr.Zero;
            var batteryStatusPointer = IntPtr.Zero;

            try
            {
                var deviceHandle = SetupDiGetClassDevs(Win32.GUID_DEVCLASS_BATTERY, Win32.DEVICE_GET_CLASS_FLAGS.DIGCF_PRESENT | Win32.DEVICE_GET_CLASS_FLAGS.DIGCF_DEVICEINTERFACE);

                var deviceInterfaceData = new Win32.SP_DEVICE_INTERFACE_DATA();
                deviceInterfaceData.CbSize = Marshal.SizeOf(deviceInterfaceData);

                SetupDiEnumDeviceInterfaces(deviceHandle, Win32.GUID_DEVCLASS_BATTERY, 0, ref deviceInterfaceData);

                deviceDataPointer = Marshal.AllocHGlobal(Win32.DEVICE_INTERFACE_BUFFER_SIZE);

                var deviceDetailData = new Win32.SP_DEVICE_INTERFACE_DETAIL_DATA { CbSize = (IntPtr.Size == 8) ? 8 : 4 + Marshal.SystemDefaultCharSize};

                SetupDiGetDeviceInterfaceDetail(deviceHandle, ref deviceInterfaceData, ref deviceDetailData, Win32.DEVICE_INTERFACE_BUFFER_SIZE);

                var batteryHandle = CreateFile(deviceDetailData.DevicePath, FileAccess.ReadWrite, FileShare.ReadWrite, FileMode.Open, Win32.FILE_ATTRIBUTES.Normal);

                var queryInformation = new Win32.BATTERY_QUERY_INFORMATION();

                DeviceIoControl(batteryHandle, Win32.IOCTL_BATTERY_QUERY_TAG, ref queryInformation.BatteryTag);

                var batteryInformation = new Win32.BATTERY_INFORMATION();
                queryInformation.InformationLevel = Win32.BATTERY_QUERY_INFORMATION_LEVEL.BatteryInformation;

                var queryInfoSize = Marshal.SizeOf(queryInformation);
                var batteryInfoSize = Marshal.SizeOf(batteryInformation);

                queryInfoPointer = Marshal.AllocHGlobal(queryInfoSize);
                Marshal.StructureToPtr(queryInformation, queryInfoPointer, false);

                batteryInfoPointer = Marshal.AllocHGlobal(batteryInfoSize);
                Marshal.StructureToPtr(batteryInformation, batteryInfoPointer, false);

                DeviceIoControl(batteryHandle, Win32.IOCTL_BATTERY_QUERY_INFORMATION, queryInfoPointer, queryInfoSize, batteryInfoPointer, batteryInfoSize);

                var updatedBatteryInformation = (Win32.BATTERY_INFORMATION)Marshal.PtrToStructure(batteryInfoPointer, typeof(Win32.BATTERY_INFORMATION));

                var batteryWaitStatus = new Win32.BATTERY_WAIT_STATUS { BatteryTag = queryInformation.BatteryTag };

                var batteryStatus = new Win32.BATTERY_STATUS();

                var waitStatusSize = Marshal.SizeOf(batteryWaitStatus);
                var batteryStatusSize = Marshal.SizeOf(batteryStatus);

                batteryWaitStatusPointer = Marshal.AllocHGlobal(waitStatusSize);
                Marshal.StructureToPtr(batteryWaitStatus, batteryWaitStatusPointer, false);

                batteryStatusPointer = Marshal.AllocHGlobal(batteryStatusSize);
                Marshal.StructureToPtr(batteryStatus, batteryStatusPointer, false);

                DeviceIoControl(batteryHandle, Win32.IOCTL_BATTERY_QUERY_STATUS, batteryWaitStatusPointer, waitStatusSize, batteryStatusPointer, batteryStatusSize);

                var updatedStatus = (Win32.BATTERY_STATUS)Marshal.PtrToStructure(batteryStatusPointer, typeof(Win32.BATTERY_STATUS));

                Win32.SetupDiDestroyDeviceInfoList(deviceHandle);

                return new BatteryInformation()
                {
                    DesignedMaxCapacity = updatedBatteryInformation.DesignedCapacity,
                    FullChargeCapacity = updatedBatteryInformation.FullChargedCapacity,
                    CurrentCapacity = updatedStatus.Capacity,
                    Voltage = updatedStatus.Voltage,
                    DischargeRate = updatedStatus.Rate
                };

            }
            finally
            {
                Marshal.FreeHGlobal(deviceDataPointer);
                Marshal.FreeHGlobal(queryInfoPointer);
                Marshal.FreeHGlobal(batteryInfoPointer);
                Marshal.FreeHGlobal(batteryStatusPointer);
                Marshal.FreeHGlobal(batteryWaitStatusPointer);
            }
        }

        private static bool DeviceIoControl(IntPtr deviceHandle, uint controlCode, ref uint output)
        {
            uint junkInput = 0;
            var retval = Win32.DeviceIoControl(deviceHandle, controlCode, ref junkInput, 0, ref output, (uint)Marshal.SizeOf(output), out uint bytesReturned, IntPtr.Zero);

            if (!retval)
            {
                var errorCode = Marshal.GetLastWin32Error();
                if (errorCode != 0)
                {
                    throw Marshal.GetExceptionForHR(errorCode);
                }
                else
                {
                    throw new Exception("DeviceIoControl call failed but Win32 didn't catch an error.");
                }
            }

            return retval;
        }

        private static bool DeviceIoControl(IntPtr deviceHandle, uint controlCode, IntPtr input, int inputSize, IntPtr output, int outputSize)
        {
            var retval = Win32.DeviceIoControl(deviceHandle, controlCode, input, (uint)inputSize, output, (uint)outputSize, out uint bytesReturned, IntPtr.Zero);

            if (retval) return retval;
            var errorCode = Marshal.GetLastWin32Error();
            if (errorCode != 0)
            {
                throw Marshal.GetExceptionForHR(errorCode);
            }
            else
            {
                throw new Exception("DeviceIoControl call failed but Win32 didn't catch an error.");
            }
        }

        private static IntPtr SetupDiGetClassDevs(Guid guid, Win32.DEVICE_GET_CLASS_FLAGS flags)
        {
            var handle = Win32.SetupDiGetClassDevs(ref guid, null, IntPtr.Zero, flags);

            if (handle == IntPtr.Zero || handle.ToInt32() == -1)
            {
                var errorCode = Marshal.GetLastWin32Error();
                if (errorCode != 0)
                {
                    throw Marshal.GetExceptionForHR(errorCode);
                }
                else
                {
                    throw new Exception("SetupDiGetClassDev call returned a bad handle.");
                }
            }
            return handle;
        }

        private static bool SetupDiEnumDeviceInterfaces(IntPtr deviceInfoSet, Guid guid, int memberIndex, ref Win32.SP_DEVICE_INTERFACE_DATA deviceInterfaceData)
        {
            var retval = Win32.SetupDiEnumDeviceInterfaces(deviceInfoSet, IntPtr.Zero, ref guid, (uint)memberIndex, ref deviceInterfaceData);

            if (!retval)
            {
                int errorCode = Marshal.GetLastWin32Error();
                if (errorCode != 0)
                {
                    if (errorCode == 259)
                    {
                        throw new Exception(
                            "SetupDeviceInfoEnumerateDeviceInterfaces ran out of batteries to enumerate.");
                    }

                    throw Marshal.GetExceptionForHR(errorCode);
                }
                else
                {
                    throw new Exception(
                        "SetupDeviceInfoEnumerateDeviceInterfaces call failed but Win32 didn't catch an error.");
                }
            }
            return retval;
        }

        private static bool SetupDiDestroyDeviceInfoList(IntPtr deviceInfoSet)
        {
            var retval = Win32.SetupDiDestroyDeviceInfoList(deviceInfoSet);

            if (!retval)
            {
                int errorCode = Marshal.GetLastWin32Error();
                if (errorCode != 0)
                {
                    throw Marshal.GetExceptionForHR(errorCode);
                }
                else
                {
                    throw new Exception("SetupDiDestroyDeviceInfoList call failed but Win32 didn't catch an error.");
                }
            }
            return retval;
        }

        private static bool SetupDiGetDeviceInterfaceDetail(IntPtr deviceInfoSet, ref Win32.SP_DEVICE_INTERFACE_DATA deviceInterfaceData, ref Win32.SP_DEVICE_INTERFACE_DETAIL_DATA deviceInterfaceDetailData, int deviceInterfaceDetailSize)
        {
            var retval = Win32.SetupDiGetDeviceInterfaceDetail(deviceInfoSet, ref deviceInterfaceData, ref deviceInterfaceDetailData, (uint)deviceInterfaceDetailSize, out uint reqSize, IntPtr.Zero);

            retval = Win32.SetupDiGetDeviceInterfaceDetail(deviceInfoSet, ref deviceInterfaceData, ref deviceInterfaceDetailData, (uint)reqSize, out reqSize, IntPtr.Zero);


            if (!retval)
            {
                int errorCode = Marshal.GetLastWin32Error();
                if (errorCode != 0)
                {
                    throw Marshal.GetExceptionForHR(errorCode);
                }
                else
                {
                    throw new Exception("SetupDiGetDeviceInterfaceDetail call failed but Win32 didn't catch an error.");
                }
            }
            return retval;
        }

        private static IntPtr CreateFile(string filename, FileAccess access, FileShare shareMode, FileMode creation, Win32.FILE_ATTRIBUTES flags)
        {
            var handle = Win32.CreateFile(filename, access, shareMode, IntPtr.Zero, creation, flags, IntPtr.Zero);

            if (handle == IntPtr.Zero || handle.ToInt32() == -1)
            {
                int errorCode = Marshal.GetLastWin32Error();
                if (errorCode != 0)
                {
                    Marshal.ThrowExceptionForHR(errorCode);
                }
                else
                {
                    throw new Exception("SetupDiGetDeviceInterfaceDetail call failed but Win32 didn't catch an error.");
                }
            }
            return handle;
        }
    }

    internal static class Win32
    {
        internal static readonly Guid GUID_DEVCLASS_BATTERY = new Guid(0x72631E54, 0x78A4, 0x11D0, 0xBC, 0xF7, 0x00, 0xAA, 0x00, 0xB7, 0xB3, 0x2A);
        internal const uint IOCTL_BATTERY_QUERY_TAG = (0x00000029 << 16) | ((int)FileAccess.Read << 14) | (0x10 << 2) | (0);
        internal const uint IOCTL_BATTERY_QUERY_INFORMATION = (0x00000029 << 16) | ((int)FileAccess.Read << 14) | (0x11 << 2) | (0);
        internal const uint IOCTL_BATTERY_QUERY_STATUS = (0x00000029 << 16) | ((int)FileAccess.Read << 14) | (0x13 << 2) | (0);

        internal const int DEVICE_INTERFACE_BUFFER_SIZE = 120;


        [DllImport("setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern IntPtr SetupDiGetClassDevs(
            ref Guid guid,
            [MarshalAs(UnmanagedType.LPTStr)] string enumerator,
            IntPtr hwnd,
            DEVICE_GET_CLASS_FLAGS flags);

        [DllImport("setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern bool SetupDiDestroyDeviceInfoList(IntPtr deviceInfoSet);

        [DllImport("setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern bool SetupDiEnumDeviceInterfaces(
            IntPtr hdevInfo,
            IntPtr devInfo,
            ref Guid guid,
            uint memberIndex,
            ref SP_DEVICE_INTERFACE_DATA devInterfaceData);

        [DllImport("setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern bool SetupDiGetDeviceInterfaceDetail(
            IntPtr hdevInfo,
            ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData,
            ref SP_DEVICE_INTERFACE_DETAIL_DATA deviceInterfaceDetailData,
            uint deviceInterfaceDetailDataSize,
            out uint requiredSize,
            IntPtr deviceInfoData);

        [DllImport("setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern bool SetupDiGetDeviceInterfaceDetail(
            IntPtr hdevInfo,
            ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData,
            IntPtr deviceInterfaceDetailData,
            uint deviceInterfaceDetailDataSize,
            out uint requiredSize,
            IntPtr deviceInfoData);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        internal static extern IntPtr CreateFile(
            string filename,
            [MarshalAs(UnmanagedType.U4)] FileAccess desiredAccess,
            [MarshalAs(UnmanagedType.U4)] FileShare shareMode,
            IntPtr securityAttributes,
            [MarshalAs(UnmanagedType.U4)] FileMode creationDisposition,
            [MarshalAs(UnmanagedType.U4)] FILE_ATTRIBUTES flags,
            IntPtr template);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        internal static extern bool DeviceIoControl(
            IntPtr handle,
            uint controlCode,
            [In] IntPtr inBuffer,
            uint inBufferSize,
            [Out] IntPtr outBuffer,
            uint outBufferSize,
            out uint bytesReturned,
            IntPtr overlapped);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        internal static extern bool DeviceIoControl(
            IntPtr handle,
            uint controlCode,
            ref uint inBuffer,
            uint inBufferSize,
            ref uint outBuffer,
            uint outBufferSize,
            out uint bytesReturned,
            IntPtr overlapped);

        [Flags]
        internal enum DEVICE_GET_CLASS_FLAGS : uint
        {
            DIGCF_DEFAULT = 0x00000001,
            DIGCF_PRESENT = 0x00000002,
            DIGCF_ALLCLASSES = 0x00000004,
            DIGCF_PROFILE = 0x00000008,
            DIGCF_DEVICEINTERFACE = 0x00000010
        }

        [Flags]
        internal enum LOCAL_MEMORY_FLAGS
        {
            LMEM_FIXED = 0x0000,
            LMEM_MOVEABLE = 0x0002,
            LMEM_NOCOMPACT = 0x0010,
            LMEM_NODISCARD = 0x0020,
            LMEM_ZEROINIT = 0x0040,
            LMEM_MODIFY = 0x0080,
            LMEM_DISCARDABLE = 0x0F00,
            LMEM_VALID_FLAGS = 0x0F72,
            LMEM_INVALID_HANDLE = 0x8000,
            LHND = (LMEM_MOVEABLE | LMEM_ZEROINIT),
            LPTR = (LMEM_FIXED | LMEM_ZEROINIT),
            NONZEROLHND = (LMEM_MOVEABLE),
            NONZEROLPTR = (LMEM_FIXED)
        }

        [Flags]
        internal enum FILE_ATTRIBUTES : uint
        {
            Readonly = 0x00000001,
            Hidden = 0x00000002,
            System = 0x00000004,
            Directory = 0x00000010,
            Archive = 0x00000020,
            Device = 0x00000040,
            Normal = 0x00000080,
            Temporary = 0x00000100,
            SparseFile = 0x00000200,
            ReparsePoint = 0x00000400,
            Compressed = 0x00000800,
            Offline = 0x00001000,
            NotContentIndexed = 0x00002000,
            Encrypted = 0x00004000,
            Write_Through = 0x80000000,
            Overlapped = 0x40000000,
            NoBuffering = 0x20000000,
            RandomAccess = 0x10000000,
            SequentialScan = 0x08000000,
            DeleteOnClose = 0x04000000,
            BackupSemantics = 0x02000000,
            PosixSemantics = 0x01000000,
            OpenReparsePoint = 0x00200000,
            OpenNoRecall = 0x00100000,
            FirstPipeInstance = 0x00080000
        }

        internal enum BATTERY_QUERY_INFORMATION_LEVEL
        {
            BatteryInformation = 0,
            BatteryGranularityInformation = 1,
            BatteryTemperature = 2,
            BatteryEstimatedTime = 3,
            BatteryDeviceName = 4,
            BatteryManufactureDate = 5,
            BatteryManufactureName = 6,
            BatteryUniqueID = 7
        }

        [Flags]
        internal enum POWER_STATE : uint
        {
            BATTERY_POWER_ONLINE = 0x00000001,
            BATTERY_DISCHARGING = 0x00000002,
            BATTERY_CHARGING = 0x00000004,
            BATTERY_CRITICAL = 0x00000008
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal struct BATTERY_INFORMATION
        {
            public int Capabilities;
            public byte Technology;

            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
            public byte[] Reserved;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 4)]
            public byte[] Chemistry;

            public int DesignedCapacity;
            public int FullChargedCapacity;
            public int DefaultAlert1;
            public int DefaultAlert2;
            public int CriticalBias;
            public int CycleCount;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal struct SP_DEVICE_INTERFACE_DETAIL_DATA
        {
            public int CbSize;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
            public string DevicePath;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct SP_DEVICE_INTERFACE_DATA
        {
            public int CbSize;
            public Guid InterfaceClassGuid;
            public int Flags;
            public UIntPtr Reserved;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal struct BATTERY_QUERY_INFORMATION
        {
            public uint BatteryTag;
            public BATTERY_QUERY_INFORMATION_LEVEL InformationLevel;
            public int AtRate;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal struct BATTERY_STATUS
        {
            public POWER_STATE PowerState;
            public uint Capacity;
            public uint Voltage;
            public int Rate;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal struct BATTERY_WAIT_STATUS
        {
            public uint BatteryTag;
            public uint Timeout;
            public POWER_STATE PowerState;
            public uint LowCapacity;
            public uint HighCapacity;
        }
    }
}