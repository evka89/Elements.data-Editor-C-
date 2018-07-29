using System;
using System.IO;
using System.Text;
using System.Windows;
using System.Collections;
using System.ComponentModel;
using System.Text.RegularExpressions;

namespace AngelicaManager.ElementsEditor
{
    public class eList
    {
        public string listName;
        public byte[] listOffset;
        public string[] elementFields;
        public string[] elementTypes;
        public object[][] elementValues;

        // return a field of an element in string representation
        public string GetValue(int ElementIndex, int FieldIndex)
        {
            if (FieldIndex > -1)
            {
                object value = elementValues[ElementIndex][FieldIndex];
                string type = elementTypes[FieldIndex];

                if (type == "bool")
                {
                    return Convert.ToString((bool)value);
                }
                if (type == "int16")
                {
                    return Convert.ToString((short)value);
                }
                if (type == "uint16")
                {
                    return Convert.ToString((ushort)value);
                }
                if (type == "int32")
                {
                    return Convert.ToString((int)value);
                }
                if (type == "uint32")
                {
                    return Convert.ToString((uint)value);
                }
                if (type == "int64")
                {
                    return Convert.ToString((long)value);
                }
                if (type == "float")
                {
                    return Convert.ToString((float)value);
                }
                if (type == "double")
                {
                    return Convert.ToString((double)value);
                }
                if (type.Contains("byte:"))
                {
                    // Convert from byte[] to Hex String
                    byte[] b = (byte[])value;
                    return BitConverter.ToString(b);
                }
                if (type.Contains("wstring:"))
                {
                    Encoding enc = Encoding.GetEncoding("Unicode");
                    return enc.GetString((byte[])value).Split('\0')[0];
                }
                if (type.Contains("string:"))
                {
                    Encoding enc = Encoding.GetEncoding("GBK");
                    return enc.GetString((byte[])value).Split('\0')[0];
                }
            }
            return "";
        }

        public void SetValue(int ElementIndex, int FieldIndex, string Value)
        {
            string type = elementTypes[FieldIndex];

            if (type == "bool")
            {
                elementValues[ElementIndex][FieldIndex] = Convert.ToBoolean(Value);
                return;
            }
            if (type == "int16")
            {
                elementValues[ElementIndex][FieldIndex] = Convert.ToInt16(Value);
                return;
            }
            if (type == "uint16")
            {
                elementValues[ElementIndex][FieldIndex] = Convert.ToUInt16(Value);
                return;
            }
            if (type == "int32")
            {
                elementValues[ElementIndex][FieldIndex] = Convert.ToInt32(Value);
                return;
            }
            if (type == "uint32")
            {
                elementValues[ElementIndex][FieldIndex] = Convert.ToUInt32(Value);
                return;
            }
            if (type == "int64")
            {
                elementValues[ElementIndex][FieldIndex] = Convert.ToInt64(Value);
                return;
            }
            if (type == "float")
            {
                elementValues[ElementIndex][FieldIndex] = Convert.ToSingle(Value);
                return;
            }
            if (type == "double")
            {
                elementValues[ElementIndex][FieldIndex] = Convert.ToDouble(Value);
                return;
            }
            if (type.Contains("byte:"))
            {
                // Convert from Hex to byte[]
                string[] hex = Value.Split(new char[] { '-' });
                byte[] b = new byte[Convert.ToInt32(type.Substring(5))];
                for (int i = 0; i < hex.Length; i++)
                    b[i] = Convert.ToByte(hex[i], 16);
                elementValues[ElementIndex][FieldIndex] = b;
                return;
            }
            if (type.Contains("wstring:"))
            {
                Encoding enc = Encoding.GetEncoding("Unicode");
                byte[] target = new byte[Convert.ToInt32(type.Substring(8))];
                byte[] source = enc.GetBytes(Value);
                if (target.Length > source.Length)
                    Array.Copy(source, target, source.Length);
                else
                    Array.Copy(source, target, target.Length);
                elementValues[ElementIndex][FieldIndex] = target;
                return;
            }
            if (type.Contains("string:"))
            {
                Encoding enc = Encoding.GetEncoding("GBK");
                byte[] target = new byte[Convert.ToInt32(type.Substring(7))];
                byte[] source = enc.GetBytes(Value);
                if (target.Length > source.Length)
                    Array.Copy(source, target, source.Length);
                else
                    Array.Copy(source, target, target.Length);
                elementValues[ElementIndex][FieldIndex] = target;
                return;
            }
            return;
        }
        // return the type of the field in string representation
        public string GetType(int FieldIndex)
        {
            if (FieldIndex > -1)
            {
                return elementTypes[FieldIndex];
            }
            return "";
        }
        // delete Item
        public void RemoveItem(int itemIndex)
        {
            object[][] newValues = new object[elementValues.Length - 1][];
            Array.Copy(elementValues, 0, newValues, 0, itemIndex);
            int length = newValues.Length - itemIndex;
            Array.Copy(elementValues, itemIndex + 1, newValues, itemIndex, newValues.Length - itemIndex);
            elementValues = newValues;
        }

        // add Item
        public void AddItem(object[] itemValues)
        {
            object[][] newValues = new object[elementValues.Length + 1][];
            Array.Resize(ref elementValues, elementValues.Length + 1);
            elementValues[elementValues.Length - 1] = itemValues;
        }
        // export item to unicode textfile
        public void ExportItem(string file, int index)
        {
            using (StreamWriter sw = new StreamWriter(file, false, Encoding.Unicode))
                for (int i = 0; i < elementTypes.Length; i++)
                    sw.WriteLine(elementFields[i] + "(" + elementTypes[i] + ")=" + GetValue(index, i));
        }
        // import item from unicode textfile
        public void ImportItem(string file, int index)
        {
            using (StreamReader sr = new StreamReader(file, Encoding.Unicode))
                for (int i = 0; i < elementTypes.Length; i++)
                    SetValue(index, i, sr.ReadLine().Split('=')[1]);
        }
        // add all new elements into the elementValues
        public ArrayList JoinElements(eList newList, int listID, bool addNew, bool backupNew, bool replaceChanged, bool backupChanged, bool removeMissing, bool backupMissing, string dirBackupNew, string dirBackupChanged, string dirBackupMissing)
        {
            object[][] newElementValues = newList.elementValues;
            string[] newElementTypes = newList.elementTypes;

            ArrayList report = new ArrayList();
            bool exists;

            // check which elements are missing (removed)
            for (int n = 0; n < elementValues.Length; n++)
            {
                //Application.DoEvents();
                exists = false;
                for (int e = 0; e < newElementValues.Length; e++)
                {
                    if (GetValue(n, 0) == newList.GetValue(e, 0))
                    {
                        exists = true;
                    }
                }
                if (!exists)
                {
                    if (dirBackupMissing != null && Directory.Exists(dirBackupMissing))
                    {
                        ExportItem(dirBackupMissing + "\\List_" + listID.ToString() + "_Item_" + GetValue(n, 0) + ".txt", n);
                    }
                    if (removeMissing)
                    {
                        report.Add("- MISSING ITEM (*removed): " + ((int)elementValues[n][0]).ToString());
                        RemoveItem(n);
                        n--;
                    }
                    else
                    {
                        report.Add("- MISSING ITEM (*not removed): " + ((int)elementValues[n][0]).ToString());
                    }
                }
            }

            for (int e = 0; e < newElementValues.Length; e++)
            {
                //Application.DoEvents();
                // check if the element with this id already exists
                exists = false;
                for (int n = 0; n < elementValues.Length; n++)
                {
                    if (GetValue(n, 0) == newList.GetValue(e, 0))
                    {
                        exists = true;
                        // check if this item is different
                        if (elementValues[n].Length != newList.elementValues[e].Length)
                        {
                            // invalid amount of values !!!
                            report.Add("<> DIFFERENT ITEM (*not replaced, invalid amount of values): " + GetValue(n, 0));
                        }
                        else
                        {
                            // compare all values of current element
                            for (int i = 0; i < elementValues[n].Length; i++)
                            {
                                if (GetValue(n, i) != newList.GetValue(e, i))
                                {
                                    if (backupChanged && Directory.Exists(dirBackupChanged))
                                    {
                                        ExportItem(dirBackupChanged + "\\List_" + listID.ToString() + "_Item_" + GetValue(n, 0) + ".txt", n);
                                    }
                                    if (replaceChanged)
                                    {
                                        report.Add("<> DIFFERENT ITEM (*replaced): " + GetValue(n, 0));
                                        elementValues[n] = newList.elementValues[e];
                                    }
                                    else
                                    {
                                        report.Add("<> DIFFERENT ITEM (*not replaced): " + GetValue(n, 0));
                                    }
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }
                if (!exists)
                {
                    if (backupNew && Directory.Exists(dirBackupNew))
                    {
                        newList.ExportItem(dirBackupNew + "\\List_" + listID.ToString() + "_Item_" + newList.GetValue(e, 0) + ".txt", e);
                    }
                    if (addNew)
                    {
                        AddItem(newElementValues[e]);
                        report.Add("+ NEW ITEM (*added): " + GetValue(elementValues.Length - 1, 0));
                    }
                    else
                    {
                        report.Add("+ NEW ITEM (*not added): " + GetValue(elementValues.Length - 1, 0));
                    }
                }
            }

            return report;
        }
    }

    public class eListCollection
    {
        public eListCollection(string game, string elFile, string cfgFile, BackgroundWorker worker)
        {
            switch (game)
            {
                case "pw":
                    Lists = LoadPW(elFile, cfgFile, worker);
                    break;
            }
        }

        FastMPPC mppc2 = new FastMPPC();

        public int SW;

        public short Version;
        public short Signature;
        public int ConversationListIndex;
        public string ConfigFile;
        public eList[] Lists;

        public void RemoveItem(int ListIndex, int ElementIndex)
        {
            Lists[ListIndex].RemoveItem(ElementIndex);
        }
        public void AddItem(int ListIndex, object[] ItemValues)
        {
            Lists[ListIndex].AddItem(ItemValues);
        }
        public string GetOffset(int ListIndex)
        {
            return BitConverter.ToString(Lists[ListIndex].listOffset);
        }

        public void SetOffset(int ListIndex, string Offset)
        {
            if (Offset != "")
            {
                // Convert from Hex to byte[]
                string[] hex = Offset.Split(new char[] { '-' });
                Lists[ListIndex].listOffset = new byte[hex.Length];
                for (int i = 0; i < hex.Length; i++)
                {
                    Lists[ListIndex].listOffset[i] = Convert.ToByte(hex[i], 16);
                }
            }
            else
            {
                Lists[ListIndex].listOffset = new byte[0];
            }
        }

        public string GetValue(int ListIndex, int ElementIndex, int FieldIndex)
        {
            return Lists[ListIndex].GetValue(ElementIndex, FieldIndex);
        }

        public void SetValue(int ListIndex, int ElementIndex, int FieldIndex, string Value)
        {
            Lists[ListIndex].SetValue(ElementIndex, FieldIndex, Value);
        }

        public string GetType(int ListIndex, int FieldIndex)
        {
            return Lists[ListIndex].GetType(FieldIndex);
        }

        private object readValue(BinaryReader br, string type)
        {
            if (type == "bool")
            {
                return br.ReadBoolean();
            }
            if (type == "int16")
            {
                return br.ReadInt16();
            }
            if (type == "uint16")
            {
                return br.ReadUInt16();
            }
            if (type == "int32")
            {
                return br.ReadInt32();
            }
            if (type == "uint32")
            {
                return br.ReadUInt32();
            }
            if (type == "int64")
            {
                return br.ReadInt64();
            }
            if (type == "float")
            {
                return br.ReadSingle();
            }
            if (type == "double")
            {
                return br.ReadDouble();
            }
            if (type.Contains("byte:"))
            {
                return br.ReadBytes(Convert.ToInt32(type.Substring(5)));
            }
            if (type.Contains("wstring:"))
            {
                return br.ReadBytes(Convert.ToInt32(type.Substring(8)));
            }
            if (type.Contains("string:"))
            {
                return br.ReadBytes(Convert.ToInt32(type.Substring(7)));
            }
            return null;
        }

        private void writeValue(BinaryWriter bw, object value, string type)
        {
            if (type == "bool")
            {
                bw.Write((bool)value);
                return;
            }
            if (type == "int16")
            {
                bw.Write((short)value);
                return;
            }
            if (type == "uint16")
            {
                bw.Write((ushort)value);
                return;
            }
            if (type == "int32")
            {
                bw.Write((int)value);
                return;
            }
            if (type == "uint32")
            {
                bw.Write((uint)value);
                return;
            }
            if (type == "int64")
            {
                bw.Write((long)value);
                return;
            }
            if (type == "float")
            {
                bw.Write((float)value);
                return;
            }
            if (type == "double")
            {
                bw.Write((double)value);
                return;
            }
            if (type.Contains("byte:"))
            {
                bw.Write((byte[])value);
                return;
            }
            if (type.Contains("wstring:"))
            {
                bw.Write((byte[])value);
                return;
            }
            if (type.Contains("string:"))
            {
                bw.Write((byte[])value);
                return;
            }
        }

        // returns an eList array with preconfigured fields from configuration file
        private eList[] loadConfiguration(string file)
        {
            StreamReader sr = new StreamReader(file);
            eList[] Li = new eList[Convert.ToInt32(sr.ReadLine())];
            try
            {
                ConversationListIndex = Convert.ToInt32(sr.ReadLine());
            }
            catch
            {
                ConversationListIndex = 58;
            }
            string line;
            for (int i = 0; i < Li.Length; i++)
            {
                //System.Windows.Forms.Application.DoEvents();

                while ((line = sr.ReadLine()) == "")
                {
                }
                Li[i] = new eList();
                Li[i].listName = line;
                Li[i].listOffset = null;
                string offset = sr.ReadLine();

                if (offset == "SW")
                {
                    if (i == 0)
                        offset = "4";

                    SW = 1;
                }

                if (offset != "AUTO")
                {
                    Li[i].listOffset = new byte[Convert.ToInt32(offset)];
                }
                Li[i].elementFields = sr.ReadLine().Split(new char[] { ';' });
                Li[i].elementTypes = sr.ReadLine().Split(new char[] { ';' });
            }
            sr.Close();

            return Li;
        }

        private Hashtable loadRules(string file)
        {
            StreamReader sr = new StreamReader(file);
            Hashtable result = new Hashtable();
            string key = "";
            string value = "";
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                //System.Windows.Forms.Application.DoEvents();

                if (line != "" && !line.StartsWith("#"))
                {
                    if (line.Contains("|"))
                    {
                        key = line.Split(new char[] { '|' })[0];
                        value = line.Split(new char[] { '|' })[1];
                    }
                    else
                    {
                        key = line;
                        value = "";
                    }
                    result.Add(key, value);
                }
            }
            sr.Close();

            return result;
        }

        public byte[] UnpackMPPC(byte[] PackChar)
        {
            return mppc2.decompress(PackChar);
        //    return (new MppcUnpacker()).UnpackMPPC(PackChar);
        }

        public byte[] PackMPPC(byte[] UnpackChar)
        {
            return mppc2.Compress2(UnpackChar);
        //    return (new FastMPPC()).Compress2(UnpackChar); ;
        }

        // only works for PW !!!
        public eList[] LoadPW(string elFile, string cfgFile, BackgroundWorker worker)
        {
            eList[] Li = new eList[0];

            // open the element file
            FileStream fs = File.OpenRead(elFile);
            BinaryReader br = new BinaryReader(fs);

            Version = br.ReadInt16();
            Signature = br.ReadInt16();
            
            string  myConfig;

            if (cfgFile == null)
                myConfig = "*_*_v" + Version;
            else
                myConfig = cfgFile;
            
            // check if a corresponding configuration file exists
            string[] configFiles = Directory.GetFiles(Environment.CurrentDirectory + "\\configs", myConfig + ".cfg"); // Application.StartupPath
            if (configFiles.Length > 0)
            {
                // configure an eList array with the configuration file
                ConfigFile = configFiles[0];
                Li = loadConfiguration(ConfigFile);

                // read the element file
                for (int l = 0; l < Li.Length; l++)
                {
                    if (worker != null)
                    {
                        int percentLoaded = (int)Math.Round((((double)l + 1) / (double)Li.Length) * 100.0);
                        worker.ReportProgress(percentLoaded, new object[] { l, Li.Length });
                    }

                    // read offset
                    if (Li[l].listOffset != null) // if (Li[l].listOffset)
                    {
                        // offset > 0
                        if (Li[l].listOffset.Length > 0)
                        {
                            Li[l].listOffset = br.ReadBytes(Li[l].listOffset.Length);
                        }
                    }
                    // autodetect offset (for list 20 & 100)
                    else
                    {
                        if (l == 20)
                        {
                            byte[] head = br.ReadBytes(4);
                            byte[] count = br.ReadBytes(4);
                            byte[] body = br.ReadBytes(BitConverter.ToInt32(count, 0));
                            byte[] tail = br.ReadBytes(4);
                            Li[l].listOffset = new byte[head.Length + count.Length + body.Length + tail.Length];
                            Array.Copy(head, 0, Li[l].listOffset, 0, head.Length);
                            Array.Copy(count, 0, Li[l].listOffset, 4, count.Length);
                            Array.Copy(body, 0, Li[l].listOffset, 8, body.Length);
                            Array.Copy(tail, 0, Li[l].listOffset, 8 + body.Length, tail.Length);
                        }
                        if (l == 100)
                        {
                            byte[] head = br.ReadBytes(4);
                            byte[] count = br.ReadBytes(4);
                            byte[] body = br.ReadBytes(BitConverter.ToInt32(count, 0));
                            Li[l].listOffset = new byte[head.Length + count.Length + body.Length];
                            Array.Copy(head, 0, Li[l].listOffset, 0, head.Length);
                            Array.Copy(count, 0, Li[l].listOffset, 4, count.Length);
                            Array.Copy(body, 0, Li[l].listOffset, 8, body.Length);
                        }
                    }

                    // read conversation list
                    if (l == ConversationListIndex && ConversationListIndex > -1)
                    {
                        // Auto detect only works for Perfect World elements.data !!!
                        if (Li[l].elementTypes[0].Contains("AUTO"))
                        {
                            if (SW != 1)
                            {
                                byte[] pattern = (Encoding.GetEncoding("GBK")).GetBytes("facedata\\");
                                long sourcePosition = br.BaseStream.Position; // Получаем позицию чтения листа
                                int listLength = -72 - pattern.Length;
                                bool run = true;
                                while (run)
                                {
                                    run = false;
                                    for (int i = 0; i < pattern.Length; i++)
                                    {
                                        listLength++;
                                        if (br.ReadByte() != pattern[i])
                                        {
                                            run = true;
                                            break;
                                        }
                                    }
                                }
                                br.BaseStream.Position = sourcePosition;
                                Li[l].elementTypes[0] = "byte:" + listLength;

                                Li[l].elementValues = new object[1][];
                                Li[l].elementValues[0] = new object[Li[l].elementTypes.Length];
                                Li[l].elementValues[0][0] = readValue(br, Li[l].elementTypes[0]);
                            }

                            if (SW == 1)
                            {
                                long MainPosition = br.BaseStream.Position;
                                int Talks = br.ReadInt32();
                                
                                for (int e = 0; e < Talks; e++)
                                {
                                    uint ID = (uint)readValue(br, "uint32");
                                    byte[] Name = (byte[])readValue(br, "wstring:128");
                                    int NUM_WINDOWS = (int)readValue(br, "int32");

                                    for (int j = 0; j < NUM_WINDOWS; j++)
                                    {
                                        uint ID_WINDOW = (uint)readValue(br, "uint32");
                                        int ID_WINDOW_PARENT = (int)readValue(br, "int32");
                                        uint TALK_TEXT_LEN = (uint)readValue(br, "uint32");
                                        byte[] TALK_TEXT = (byte[])readValue(br, ("wstring:" + TALK_TEXT_LEN * 2));
                                        int NUM_OPTION = (int)readValue(br, "int32");

                                        for (int k = 0; k < NUM_OPTION; k++)
                                        {
                                            int ID_OPTION = (int)readValue(br, "int32");
                                            byte[] TEXT_OPTION = (byte[])readValue(br, "wstring:128");
                                            uint PARAM_OPTION = (uint)readValue(br, "uint32");
                                        }
                                    }
                                    int TALK_PROC_TYPE = (int)readValue(br, "int32");
                                    uint ID_PATH = (uint)readValue(br, "uint32");

                                }

                                int AfterTalkLen = br.ReadInt32();
                                byte[] AfterTalk = br.ReadBytes(AfterTalkLen);

                                if (AfterTalkLen != AfterTalk.Length)
                                    MessageBox.Show("В данном Элике проблема с окончанием диалогов, \nпосле сохранения может быть не рабочим!");

                                long listLength = fs.Position - MainPosition;

                                br.BaseStream.Position = MainPosition;
                                Li[l].elementTypes[0] = "byte:" + listLength;

                                Li[l].elementValues = new object[1][];
                                Li[l].elementValues[0] = new object[Li[l].elementTypes.Length];
                                Li[l].elementValues[0][0] = readValue(br, Li[l].elementTypes[0]);
                            }
                        }
                    }
                    // read lists
                    else
                    {
                        try
                        {
                            Li[l].elementValues = new object[br.ReadInt32()][];
                        }
                        catch
                        {
                            return Li;
                        }

                        ushort[] FullSize = new ushort[Li[l].elementValues.Length];
                        uint[] id = new uint[Li[l].elementValues.Length];

                        uint sum_str_len;
                        long el_size = 0;

                        MemoryStream fsList = new MemoryStream();
                        BinaryWriter bwList = new BinaryWriter(fsList);

                        if (SW == 1 && l != ConversationListIndex)
                        {
                            ushort nameSize;

                            for (int f = 0; f < Li[l].elementValues.Length; f++)
                            {
                                id[f] = br.ReadUInt32();
                                nameSize = br.ReadUInt16();
                                FullSize[f] = nameSize;
                            }

                            if (Li[l].elementValues.Length > 0)
                            {
                                sum_str_len = br.ReadUInt32();

                                for (int i = 0; i < Li[l].elementValues.Length; i++)
                                {
                                    bwList.Write(UnpackMPPC(br.ReadBytes(FullSize[i])));
                                    if (i == 0)
                                        el_size = fsList.Position;
                                }
                            }
                        }

                        fsList.Position = 0;
                        BinaryReader brSW = new BinaryReader(fsList);

                        // go through all elements of a list
                        // sdfsdf asdl;zcmv,bl
                        for (int e = 0; e < Li[l].elementValues.Length; e++)
                        {
                            Li[l].elementValues[e] = new object[Li[l].elementTypes.Length];

                            // go through all fields of an element
                            for (int f = 0; f < Li[l].elementValues[e].Length; f++)
                            {
                                if (SW == 1)
                                {
                                    Li[l].elementValues[e][f] = readValue(brSW, Li[l].elementTypes[f]);
                                    if (e == 0 && f == Li[l].elementValues[e].Length - 1 && fsList.Position != el_size)
                                        MessageBox.Show($"Лист {l}({Li[l].listName})\nТекущий размер: {fsList.Position} байт\nИсходный размер: {el_size} байт\nКол-во предметов: {Li[l].elementValues.Length}");
                                }
                                if (SW != 1)
                                    Li[l].elementValues[e][f] = readValue(br, Li[l].elementTypes[f]);
                            }
                        }

                        brSW.Close();
                        bwList.Close();
                        fsList.Flush();
                        fsList.Close();
                    }
                }
            }
            else
            {
                MessageBox.Show("No corressponding configuration file found!\nVersion: " + Version + "\nPattern: " + "configs\\PW_*_v" + Version + ".cfg");
            }

            br.Close();
            fs.Close();

            return Li;
        }

        public void SavePW(string elFile, BackgroundWorker worker)
        {
            if (File.Exists(elFile))
                File.Delete(elFile);

            using (BinaryWriter bw = new BinaryWriter(new FileStream(elFile, FileMode.Create, FileAccess.Write)))
            {

                bw.Write(Version);
                bw.Write(Signature);

                // go through all lists
                for (int l = 0; l < Lists.Length; l++)
                {
                    if (worker != null)
                    {
                        int percentLoaded = (int)Math.Round((((double)l + 1) / (double)Lists.Length) * 100.0);
                        worker.ReportProgress(percentLoaded);
                    }
                    //System.Windows.Forms.Application.DoEvents();

                    if (Lists[l].listOffset.Length > 0)
                    {
                        bw.Write(Lists[l].listOffset);
                    }

                    if (l != ConversationListIndex)
                    {
                        bw.Write(Lists[l].elementValues.Length);
                    }

                    if (SW == 1 && l != ConversationListIndex && Lists[l].elementValues.Length != 0)
                    {
                        uint ListSize = 0;
                        ushort elSize = 0;

                        MemoryStream fsSW = new MemoryStream();
                        BinaryWriter bwSW = new BinaryWriter(fsSW);
                        BinaryReader brSW = new BinaryReader(fsSW);

                        for (int i = 0; i < Lists[l].elementValues.Length; i++)
                        {
                            writeValue(bw, Lists[l].elementValues[i][0], Lists[l].elementTypes[0]);

                            //Запись используемых байт
                            MemoryStream fsSave = new MemoryStream();
                            BinaryWriter bwSave = new BinaryWriter(fsSave);
                            BinaryReader brSave = new BinaryReader(fsSave);

                            for (int j = 0; j < Lists[l].elementValues[i].Length; j++)
                                writeValue(bwSave, Lists[l].elementValues[i][j], Lists[l].elementTypes[j]);

                            int size = (int)fsSave.Position;
                            fsSave.Position = 0;

                            byte[] SaveValue = brSave.ReadBytes(size);

                            SaveValue = PackMPPC(SaveValue);

                            bwSave.Close();
                            brSave.Close();
                            fsSave.Flush();
                            fsSave.Close();

                            elSize = Convert.ToUInt16(SaveValue.Length);
                            bwSW.Write(SaveValue);

                            bw.Write((ushort)elSize);
                            ListSize += elSize;
                        }

                        bw.Write(ListSize);
                        fsSW.Position = 0;
                        bw.Write(brSW.ReadBytes((int)ListSize));

                        bwSW.Close();
                        brSW.Close();
                        fsSW.Flush();
                        fsSW.Close();
                    }

                    if (SW != 1 || (SW == 1 && l == ConversationListIndex))
                    {
                        // Проход через все элементы в листу
                        for (int e = 0; e < Lists[l].elementValues.Length; e++)
                        {
                            // Проход через все строки элемента
                            for (int f = 0; f < Lists[l].elementValues[e].Length; f++)
                            {
                                writeValue(bw, Lists[l].elementValues[e][f], Lists[l].elementTypes[f]);
                            }
                        }
                    }
                }

                if (Version == 80 && Lists.Length == 112) bw.Write(0);
            }

        //    MessageBox.Show("Успешно сохранено!");
        }

        public void Export(string RulesFile, string TargetFile)
        {
            // Load the rules
            Hashtable rules = loadRules(RulesFile);

            if (File.Exists(TargetFile))
                File.Delete(TargetFile);

            using (BinaryWriter bw = new BinaryWriter(new FileStream(TargetFile, FileMode.Create, FileAccess.Write)))
            {
                if (rules.ContainsKey("SETVERSION"))
                {
                    bw.Write(Convert.ToInt16((string)rules["SETVERSION"]));
                }
                else
                {
                    MessageBox.Show("Rules file is missing parameter\n\nSETVERSION:", "Export Failed");
                    return;
                }

                if (rules.ContainsKey("SETSIGNATURE"))
                {
                    bw.Write(Convert.ToInt16((string)rules["SETSIGNATURE"]));
                }
                else
                {
                    MessageBox.Show("Rules file is missing parameter\n\nSETSIGNATURE:", "Export Failed");
                    return;
                }

                // go through all lists
                for (int l = 0; l < Lists.Length; l++)
                {
                    if (!rules.ContainsKey("REMOVELIST:" + l))
                    {
                        if (rules.ContainsKey("REPLACEOFFSET:" + l))
                        {
                            // Convert from Hex to byte[]
                            string[] hex = ((string)rules["REPLACEOFFSET:" + l]).Split(new char[] { '-', ' ' });
                            if (hex.Length > 1)
                            {
                                byte[] b = new byte[hex.Length];
                                for (int i = 0; i < hex.Length; i++)
                                    b[i] = Convert.ToByte(hex[i], 16);
                                if (b.Length > 0)
                                    bw.Write(b);
                            }
                        }
                        else
                        {
                            if (Lists[l].listOffset.Length > 0)
                            {
                                bw.Write(Lists[l].listOffset);
                            }
                        }

                        if (l != ConversationListIndex)
                        {
                            bw.Write(Lists[l].elementValues.Length);
                        }

                        // go through all elements of a list
                        for (int e = 0; e < Lists[l].elementValues.Length; e++)
                        {
                            // go through all fields of an element
                            for (int f = 0; f < Lists[l].elementValues[e].Length; f++)
                            {
                                //System.Windows.Forms.Application.DoEvents();

                                if (!rules.ContainsKey("REMOVEVALUE:" + l + ":" + f))
                                {
                                    writeValue(bw, Lists[l].elementValues[e][f], Lists[l].elementTypes[f]);
                                }
                            }
                        }
                    }
                }
            }
        }
        
    }

    public class eChoice
    {
        public int Control;
        public byte[] ChoiceText;
        public int PARAM_OPTION;

        public string GetText()
        {
            Encoding enc = Encoding.GetEncoding("Unicode");
            return enc.GetString(ChoiceText);
        }

        public void SetText(string Value)
        {
            Encoding enc = Encoding.GetEncoding("Unicode");
            byte[] target = new byte[128];
            byte[] source = enc.GetBytes(Value);
            if (target.Length > source.Length)
                Array.Copy(source, target, source.Length);
            else
                Array.Copy(source, target, target.Length);
            ChoiceText = target;
        }
    }

    public class eQuestion
    {
        public int QuestionID;
        public int Control;
        public int QuestionTextLength;
        public byte[] QuestionText;

        public string GetText()
        {
            Encoding enc = Encoding.GetEncoding("Unicode");
            return enc.GetString(QuestionText);
        }

        public void SetText(string Value)
        {
            Encoding enc = Encoding.GetEncoding("Unicode");
            QuestionText = enc.GetBytes(Value + (char)0);
            QuestionTextLength = QuestionText.Length / 2;
        }

        public int ChoiceQount;
        public eChoice[] Choices;
    }

    public class eDialog
    {
        public int DialogID;
        public byte[] DialogName;

        public string GetText()
        {
            Encoding enc = Encoding.GetEncoding("Unicode");
            return enc.GetString(DialogName);
        }

        public void SetText(string Value)
        {
            Encoding enc = Encoding.GetEncoding("Unicode");
            byte[] target = new byte[128];
            byte[] source = enc.GetBytes(Value);
            if (target.Length > source.Length)
                Array.Copy(source, target, source.Length);
            else
                Array.Copy(source, target, target.Length);
            DialogName = target;
        }
        public int QuestionCount;
        public eQuestion[] Questions;

        public int TALK_PROC_TYPE;
        public int ID_PATH;
    }

    public class eListConversation
    {
        public eListConversation(byte[] Bytes)
        {
            using (BinaryReader br = new BinaryReader(new MemoryStream(Bytes)))
            {
                DialogCount = br.ReadInt32();
                Dialogs = new eDialog[DialogCount];
                for (int d = 0; d < DialogCount; d++)
                {
                    Dialogs[d] = new eDialog();
                    Dialogs[d].DialogID = br.ReadInt32();
                    Dialogs[d].DialogName = br.ReadBytes(128);
                    Dialogs[d].QuestionCount = br.ReadInt32();
                    Dialogs[d].Questions = new eQuestion[Dialogs[d].QuestionCount];
                    for (int q = 0; q < Dialogs[d].QuestionCount; q++)
                    {
                        Dialogs[d].Questions[q] = new eQuestion();
                        Dialogs[d].Questions[q].QuestionID = br.ReadInt32();
                        Dialogs[d].Questions[q].Control = br.ReadInt32();
                        Dialogs[d].Questions[q].QuestionTextLength = br.ReadInt32();
                        Dialogs[d].Questions[q].QuestionText = br.ReadBytes(2 * Dialogs[d].Questions[q].QuestionTextLength);
                        Dialogs[d].Questions[q].ChoiceQount = br.ReadInt32();
                        Dialogs[d].Questions[q].Choices = new eChoice[Dialogs[d].Questions[q].ChoiceQount];
                        for (int c = 0; c < Dialogs[d].Questions[q].ChoiceQount; c++)
                        {
                            Dialogs[d].Questions[q].Choices[c] = new eChoice();
                            Dialogs[d].Questions[q].Choices[c].Control = br.ReadInt32();
                            Dialogs[d].Questions[q].Choices[c].ChoiceText = br.ReadBytes(128);
                            Dialogs[d].Questions[q].Choices[c].PARAM_OPTION = br.ReadInt32();
                        }
                    }
                }
            }
        }
        void Dispose()
        {
            DialogCount = 0;
            //delete Count;
            Dialogs = null;
        }

        public int DialogCount;
        public eDialog[] Dialogs;

        public byte[] GetBytes()
        {
            MemoryStream ms = new MemoryStream(DialogCount);
            BinaryWriter bw = new BinaryWriter(ms);

            bw.Write(DialogCount);
            for (int d = 0; d < DialogCount; d++)
            {
                bw.Write(Dialogs[d].DialogID);
                bw.Write(Dialogs[d].DialogName);
                bw.Write(Dialogs[d].QuestionCount);
                for (int q = 0; q < Dialogs[d].QuestionCount; q++)
                {
                    bw.Write(Dialogs[d].Questions[q].QuestionID);
                    bw.Write(Dialogs[d].Questions[q].Control);
                    bw.Write(Dialogs[d].Questions[q].QuestionTextLength);
                    bw.Write(Dialogs[d].Questions[q].QuestionText);
                    bw.Write(Dialogs[d].Questions[q].ChoiceQount);
                    for (int c = 0; c < Dialogs[d].Questions[q].ChoiceQount; c++)
                    {
                        bw.Write(Dialogs[d].Questions[q].Choices[c].Control);
                        bw.Write(Dialogs[d].Questions[q].Choices[c].ChoiceText);
                        bw.Write(Dialogs[d].Questions[q].Choices[c].PARAM_OPTION);
                    }
                }
            }

            byte[] result = ms.ToArray();
            bw.Close();
            ms.Close();
            return result;
        }
    }

    public class eListConversationSW
    {
        public eListConversationSW(byte[] Bytes)
        {
            MemoryStream ms = new MemoryStream(Bytes);
            BinaryReader br = new BinaryReader(ms);
            DialogCount = br.ReadInt32(); // Количество диалогов
            Dialogs = new eDialog[DialogCount];
            for (int d = 0; d < DialogCount; d++) // Структура диалогов
            {
                Dialogs[d] = new eDialog();
                Dialogs[d].DialogID = br.ReadInt32(); // ID
                Dialogs[d].DialogName = br.ReadBytes(128); // Name
                Dialogs[d].QuestionCount = br.ReadInt32(); // NUM_WINDOWS
                Dialogs[d].Questions = new eQuestion[Dialogs[d].QuestionCount];
                for (int q = 0; q < Dialogs[d].QuestionCount; q++)
                {
                    Dialogs[d].Questions[q] = new eQuestion();
                    Dialogs[d].Questions[q].QuestionID = br.ReadInt32(); // ID_WINDOW
                    Dialogs[d].Questions[q].Control = br.ReadInt32(); // ID_WINDOW_PARENT
                    Dialogs[d].Questions[q].QuestionTextLength = br.ReadInt32(); // TALK_TEXT_LEN
                    Dialogs[d].Questions[q].QuestionText = br.ReadBytes(2 * Dialogs[d].Questions[q].QuestionTextLength); // TALK_TEXT (TALK_TEXT_LEN * 2)
                    Dialogs[d].Questions[q].ChoiceQount = br.ReadInt32(); // NUM_OPTION
                    Dialogs[d].Questions[q].Choices = new eChoice[Dialogs[d].Questions[q].ChoiceQount];
                    for (int c = 0; c < Dialogs[d].Questions[q].ChoiceQount; c++)
                    {
                        Dialogs[d].Questions[q].Choices[c] = new eChoice();
                        Dialogs[d].Questions[q].Choices[c].Control = br.ReadInt32(); // ID_OPTION
                        Dialogs[d].Questions[q].Choices[c].ChoiceText = br.ReadBytes(128); // TEXT_OPTION
                        Dialogs[d].Questions[q].Choices[c].PARAM_OPTION = br.ReadInt32(); // PARAM_OPTION
                    }
                }
                Dialogs[d].TALK_PROC_TYPE = br.ReadInt32(); // TALK_PROC_TYPE
                Dialogs[d].ID_PATH = br.ReadInt32(); // ID_PATH
            }
            AfterTalkLen = br.ReadInt32();
            AfterTalk = br.ReadBytes(AfterTalkLen);

            br.Close();
            ms.Close();
        }
        void Dispose()
        {
            DialogCount = 0;
            //delete Count;
            Dialogs = null;
        }

        public int DialogCount;
        public int AfterTalkLen;
        public byte[] AfterTalk;
        public eDialog[] Dialogs;
        public byte[] GetBytes()
        {
            MemoryStream ms = new MemoryStream(DialogCount);
            BinaryWriter bw = new BinaryWriter(ms);

            bw.Write(DialogCount);
            for (int d = 0; d < DialogCount; d++)
            {
                bw.Write(Dialogs[d].DialogID);
                bw.Write(Dialogs[d].DialogName);
                bw.Write(Dialogs[d].QuestionCount);
                for (int q = 0; q < Dialogs[d].QuestionCount; q++)
                {
                    bw.Write(Dialogs[d].Questions[q].QuestionID);
                    bw.Write(Dialogs[d].Questions[q].Control);
                    bw.Write(Dialogs[d].Questions[q].QuestionTextLength);
                    bw.Write(Dialogs[d].Questions[q].QuestionText);
                    bw.Write(Dialogs[d].Questions[q].ChoiceQount);
                    for (int c = 0; c < Dialogs[d].Questions[q].ChoiceQount; c++)
                    {
                        bw.Write(Dialogs[d].Questions[q].Choices[c].Control);
                        bw.Write(Dialogs[d].Questions[q].Choices[c].ChoiceText);
                        bw.Write(Dialogs[d].Questions[q].Choices[c].PARAM_OPTION);
                    }
                }
                bw.Write(Dialogs[d].TALK_PROC_TYPE);
                bw.Write(Dialogs[d].ID_PATH);
            }
            bw.Write(AfterTalkLen);
            bw.Write(AfterTalk);

            byte[] result = ms.ToArray();
            bw.Close();
            ms.Close();
            return result;
        }
    }
}