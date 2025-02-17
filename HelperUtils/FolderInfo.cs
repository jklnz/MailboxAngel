﻿using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace HelperUtils
{
    /// <summary>
    /// Class for holding folder information as part of a list of folder history
    /// Implements the ILimitedQueueItem interface
    /// </summary>
    public class FolderInfo : HistoryListItemBase
    {
        const string NAME_ERROR = "<Name Error>";
        public FolderInfo(MAPIFolder folder)
        {
            this._folder = folder;
            _entryID = folder.EntryID;
            _storeID = folder.StoreID;
        }
        public FolderInfo(string entryID, string storeID)
        {
            _entryID = entryID;
            _storeID = storeID;
            _folder = null;
        }


        public override bool Active
        { get { return _folder != null; } }

        public override string UniqueID
        { get { return _folder.EntryID; } }

        private MAPIFolder _folder;
        public MAPIFolder Folder
        {
            get { return _folder; }
        }

        public new bool Persist
        {
            get { return _persist; }
            set { _persist = value; }
        }

        public new bool Avoid
        {
            get { return _avoid; }
            set { _avoid = value; }
        }

        private string _entryID;
        public string EntryID
        {
            get { return _entryID; }
        }

        private string _storeID;
        public string StoreID
        {
            get { return _storeID; }
        }

        public string Path
        {
            get
            {
                try
                {
                    return _folder.FullFolderPath;
                }
                catch (COMException)
                {

                    return NAME_ERROR;
                }
            }
        }

        public string Name
        {
            get
            {
                try
                {
                    int nameStart = _folder.FullFolderPath.LastIndexOf(@"\");
                    if (nameStart > 0)
                    {
                        return string.Format("{0} ({1})", _folder.FullFolderPath.Substring(nameStart + 1), TrimPath(_folder.FullFolderPath.Substring(0, nameStart)));
                    }
                    else
                        return _folder.FullFolderPath;

                }
                catch (COMException)
                {

                    return NAME_ERROR;
                }
            }
        }

        public override string ToString()
        {
            try
            {
                return _folder.FullFolderPath;
            }
            catch (System.Exception)
            {

                return NAME_ERROR;
            }

        }

        private string TrimPath(string source)
        {
            const int LENGTH = 25;
            string[] arr = source.Split('\\');
            for (int i = 0; i < arr.Length; i++)
            {
                if (arr[i].Length > LENGTH)
                    arr[i] = arr[i].Substring(0, LENGTH - 1) + "...";
            }
            return String.Join(@"\", arr);

        }
    }
}
