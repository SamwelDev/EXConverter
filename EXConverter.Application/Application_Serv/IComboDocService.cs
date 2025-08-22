using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EXConverter.Application.Application_Serv
{
    public  interface IComboDocService
    {
        //***merging of docs
        byte[] MergeDocuments(List<byte[]> documents);
        //*****slpiting of documents
        List<byte[]> SplitDocument(byte[] documentBytes, List<int> pageIndices);
    }
}
