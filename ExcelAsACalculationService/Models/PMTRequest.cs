//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelAsACalculationService.Models
{
    public class PMTRequest : IExcelRequest
    {
        public decimal Rate { get; set; }
        public decimal Nper { get; set; }
        public decimal Pv { get; set; }
    }
}
