using System;
using System.Collections.Generic;
using System.Text;

namespace CorrectContraAccountLogicDLL
{
    public interface ICompany
    {
        SAPbobsCOM.Company Company { get; set; }
    }
}
