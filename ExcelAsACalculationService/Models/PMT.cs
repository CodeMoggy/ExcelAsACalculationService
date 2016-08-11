//Copyright (c) CodeMoggy. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using Newtonsoft.Json.Serialization;
using Newtonsoft.Json;

namespace ExcelAsACalculationService.Models
{
    public class PMT
    {
        [Required]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:C}")]
        [Display(Name = "Interest Rate (%)")]
        public decimal InterestRate { get; set; }

        [Required]
        [Display(Name = "Loan Period (months)")]
        public int NumberOfMonths { get; set; }

        [Required]
        [Display(Name = "Loan Amount")]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:C}")]
        public decimal LoanAmount { get; set; }

        [Display(Name = "Amount per Month")]
        public decimal MonthlyPaymentAmount { get; set; }

        public PMT(
            decimal loanAmount,
            decimal interestRate,
            int numberofMonths,
            decimal monthlyPaymentAmount)
        {
            InterestRate = interestRate;
            NumberOfMonths = numberofMonths;
            MonthlyPaymentAmount = monthlyPaymentAmount;
            LoanAmount = loanAmount;
        }

        public PMT() { }
    }
}