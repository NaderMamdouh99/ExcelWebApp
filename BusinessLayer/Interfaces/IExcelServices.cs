using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.AspNetCore.Http;

namespace BusinessLayer.Interfaces
{
    public interface IExcelServices
    {
        byte[] ProcessExcelFile(IFormFile filename);
    }
}
