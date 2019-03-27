using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using HGM.Hotbird64.Vlmcs;

namespace HGM.Hotbird64.LicenseManager
{
    [StructLayout(LayoutKind.Sequential)]
    public unsafe struct VlmcsdData
    {
        public KmsGuid Guid;
        public ulong NameOffset;
        public byte AppIndex;
        public byte KmsIndex;
        public byte ProtocolVersion;
        public byte NCountPolicy;
        public byte IsRetail;
        public byte IsPreview;
        public byte EPidIndex;
        public byte Reserved;

        public static ulong Size => (uint)sizeof(VlmcsdData);
    }

    [StructLayout(LayoutKind.Sequential)]
    public unsafe struct CsvlkData
    {
        public ulong EPidOffset;
        public uint GroupId;
        public uint MinKeyId;
        public uint MaxKeyId;
        public byte MinActiveClients;
        public fixed byte Reserved[3];

        public static ulong Size => (uint)sizeof(CsvlkData);
    }

    [Flags]
    public enum VlmcsdOption
    {
        None = 0,
        UseNdr64 = 1 << 0,
    }

    [StructLayout(LayoutKind.Sequential)]
    public unsafe struct VlmcsdHeader
    {
        public fixed byte Magic[4];
        public uint Version;
        public byte CsvlkCount;
        public byte Flags;
        public fixed byte Reserved[2];
        public uint AppItemCount;
        public uint KmsItemCount;
        public uint SkuItemCount;
        public ulong AppItemOffset;
        public ulong KmsItemOffset;
        public ulong SkuItemOffset;

        public static ulong Size => (uint)sizeof(VlmcsdHeader);

        public MemoryStream WriteData(bool includeAppItems, bool includeKmsItems, bool includeSkuItems, bool noText, bool includeBetaSkuItem)
        {
            fixed (byte* m = Magic)
            {
                m[0] = 0x4B;
                m[1] = 0x4D;
                m[2] = 0x44;
                m[3] = 0x00;
            }

            var exportedCsvlks = KmsLists.CsvlkItemList.Where(c => c.VlmcsdIndex >= 0).OrderBy(c => c.VlmcsdIndex).ToArray() as IReadOnlyList<CsvlkItem>;
            CsvlkCount = (byte)exportedCsvlks.Count;

            try
            {
                var ePid = new EPid(exportedCsvlks[0].EPid);
                Flags |= (byte)(ePid.OsBuild > 7601 ? VlmcsdOption.UseNdr64 : VlmcsdOption.None);
            }
            catch
            {
                // No valid host build
            }

            var kmsItemList = includeKmsItems ? KmsLists.KmsItemList.ToArray() : KmsLists.KmsItemList.Where(k => !k.CanMapToDefaultCsvlk).ToArray();

            var skuItemList = KmsLists.SkuItemList.Where(s => includeBetaSkuItem || !s.KmsItem.IsPreview).ToArray();
            var csvlkDataSize = CsvlkData.Size * CsvlkCount;

            Version = KmsLists.KmsData.Version.Full;
            AppItemOffset = Size + csvlkDataSize;
            AppItemCount = includeAppItems ? (uint)KmsLists.AppItemList.Count : 0;
            KmsItemOffset = AppItemOffset + AppItemCount * VlmcsdData.Size;
            KmsItemCount = (uint)kmsItemList.Length;
            SkuItemOffset = KmsItemOffset + KmsItemCount * VlmcsdData.Size;
            SkuItemCount = includeSkuItems && includeKmsItems ? (uint)skuItemList.Length : 0;

            var idData = new VlmcsdData[AppItemCount + KmsItemCount + SkuItemCount];
            var textData = new VlmcsdDataText[idData.Length];
            var currentText = (ulong)idData.Length * VlmcsdData.Size + Size + csvlkDataSize;
            var ePids = new VlmcsdDataText[CsvlkCount];
            var iniFileNames = new VlmcsdDataText[CsvlkCount];
            var csvlkDataList = new CsvlkData[CsvlkCount];

            ePids[0] = new VlmcsdDataText(exportedCsvlks[0].EPid, currentText);


            for (var i = 0; i < CsvlkCount; i++)
            {
                if (i != 0)
                {
                    ePids[i] = new VlmcsdDataText(exportedCsvlks[i].EPid, iniFileNames[i - 1].OffsetNext);
                }

                iniFileNames[i] = new VlmcsdDataText(exportedCsvlks[i].IniFileName, ePids[i].OffsetNext);

                csvlkDataList[i] = new CsvlkData
                {
                    EPidOffset = ePids[i].Offset,
                    GroupId = (uint)exportedCsvlks[i].GroupId,
                    MinKeyId = (uint)exportedCsvlks[i].MinKeyId,
                    MaxKeyId = (uint)exportedCsvlks[i].MaxKeyId,
                    MinActiveClients = 0,
                };
            }

            currentText = iniFileNames[CsvlkCount - 1].OffsetNext;

            var unknownText = new VlmcsdDataText("Unknown", currentText);

            if (noText)
            {
                currentText = unknownText.OffsetNext;
            }

            var index = 0;

            if (includeAppItems)
            {
                foreach (var appItem in KmsLists.AppItemList)
                {
                    var minActiveClients = (byte)appItem.MinActiveClients;

                    if (minActiveClients < 1)
                    {
                        minActiveClients = (byte)(appItem.KmsItems.Select(k => k.NCountPolicy).Max() << 1);
                    }

                    idData[index].Guid = appItem.Guid;
                    idData[index].EPidIndex = (byte)appItem.VlmcsdIndex;
                    //TODO: Find out if client count is really per AppItem
                    idData[index].NCountPolicy = minActiveClients;
                    index = WriteText(noText, idData, index, unknownText, textData, appItem, ref currentText);
                }
            }

            foreach (var kmsItem in kmsItemList)
            {
                idData[index].Guid = kmsItem.Guid;
                idData[index].IsPreview = (byte)(kmsItem.IsPreview ? 1 : 0);
                idData[index].IsRetail = (byte)(kmsItem.IsRetail ? 1 : 0);
                idData[index].NCountPolicy = (byte)kmsItem.NCountPolicy;
                idData[index].ProtocolVersion = (byte)kmsItem.DefaultKmsProtocol.Major;
                idData[index].AppIndex = (byte)(includeAppItems ? KmsLists.AppItemList.IndexOf(kmsItem.App) : 0);
                var csvlk = exportedCsvlks.SingleOrDefault(c => c.Activates[kmsItem.Guid] != null);

                if (csvlk != null)
                {
                    idData[index].EPidIndex = (byte)csvlk.VlmcsdIndex;
                }

                index = WriteText(noText, idData, index, unknownText, textData, kmsItem, ref currentText);
            }

            if (includeKmsItems && includeSkuItems)
            {
                foreach (var skuItem in skuItemList)
                {
                    idData[index].Guid = skuItem.Guid;
                    idData[index].KmsIndex = (byte)KmsLists.KmsItemList.IndexOf(skuItem.KmsItem);
                    idData[index].AppIndex = (byte)KmsLists.AppItemList.IndexOf(skuItem.KmsItem.App);
                    idData[index].ProtocolVersion = (byte)skuItem.KmsItem.DefaultKmsProtocol.Major;
                    idData[index].NCountPolicy = (byte)skuItem.KmsItem.NCountPolicy;
                    idData[index].IsPreview = (byte)(skuItem.KmsItem.IsPreview ? 1 : 0);
                    idData[index].IsRetail = (byte)(skuItem.KmsItem.IsRetail ? 1 : 0);
                    var csvlk = exportedCsvlks.SingleOrDefault(c => c.Activates[skuItem.KmsItem.Guid] != null);

                    if (csvlk != null)
                    {
                        idData[index].EPidIndex = (byte)csvlk.VlmcsdIndex;
                    }

                    index = WriteText(noText, idData, index, unknownText, textData, skuItem, ref currentText);
                }
            }

            var stream = new MemoryStream();
            stream.SetLength((long)currentText);
            stream.Seek(0, SeekOrigin.Begin);

            fixed (VlmcsdHeader* b = &this)
            {
                for (ulong i = 0; i < Size; i++)
                {
                    stream.WriteByte(((byte*)b)[i]);
                }
            }

            for (var i = 0; i < CsvlkCount; i++)
            {
                fixed (CsvlkData* b = &csvlkDataList[i])
                {
                    for (ulong j = 0; j < CsvlkData.Size; j++)
                    {
                        stream.WriteByte(((byte*)b)[j]);
                    }
                }
            }

            // ReSharper disable once ForCanBeConvertedToForeach
            for (var i = 0; i < idData.Length; i++)
            {
                fixed (VlmcsdData* b = &idData[i])
                {
                    for (ulong j = 0; j < VlmcsdData.Size; j++)
                    {
                        stream.WriteByte(((byte*)b)[j]);
                    }
                }
            }

            for (var i = 0; i < ePids.Length; i++)
            {
                stream.Write(ePids[i].Data, 0, ePids[i].Data.Length);
                stream.Write(iniFileNames[i].Data, 0, iniFileNames[i].Data.Length);
            }

            if (noText)
            {
                stream.Write(unknownText.Data, 0, unknownText.Data.Length);
            }
            else
            {
                foreach (var text in textData)
                {
                    stream.Write(text.Data, 0, text.Data.Length);
                }
            }

            return stream;
        }

        private static int WriteText(bool noText, VlmcsdData[] idData, int index, VlmcsdDataText unknownText, IList<VlmcsdDataText> textData, KmsProduct item, ref ulong currentText)
        {
            if (noText)
            {
                idData[index].NameOffset = unknownText.Offset;
            }
            else
            {
                textData[index] = new VlmcsdDataText(item.DisplayName, currentText);
                idData[index].NameOffset = currentText;
                currentText = textData[index].OffsetNext;
            }

            return index + 1;
        }
    }

    public class VlmcsdDataText
    {
        public byte[] Data;
        public ulong Offset;
        public ulong OffsetNext => Offset + (ulong)Data.LongLength;

        public VlmcsdDataText(string text, ulong offset)
        {
            Offset = offset;
            Data = new UTF8Encoding(false, false).GetBytes(text + (char)0);
        }
    }
}