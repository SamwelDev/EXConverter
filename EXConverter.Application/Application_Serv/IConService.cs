using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EXConverter.Application.Application_Serv
{
    public interface IConService
    {
        byte[] ConvertExcelToPdf(Stream excelStream);
    }
}
