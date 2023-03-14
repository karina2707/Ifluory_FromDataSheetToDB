using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LabFourSQL
{
    class Company
    {
        private bool Bankrupt;
        private double AccountsReceivableTurnover;
        private double OperatingGrossMargin;
        private double CashCurrentLiability;

        public Company(bool bankrupt, double accountsReceivableTurnover, 
                                  double operatingGrossMargin, double cashCurrentLiability) {

            this.Bankrupt = bankrupt;
            this.AccountsReceivableTurnover = accountsReceivableTurnover;
            this.OperatingGrossMargin = operatingGrossMargin;
            this.CashCurrentLiability = cashCurrentLiability;
        }

        public bool getBankrupt()
        {
            return this.Bankrupt;
        }
        public double getAccountsReceivableTurnover()
        {
            return this.AccountsReceivableTurnover;
        }
        public double getOperatingGrossMargin()
        {
            return this.OperatingGrossMargin;
        }
        public double getCashCurrentLiability()
        {
            return this.CashCurrentLiability;
        }
        public void setBankrupt(bool newValue)
        {
            this.Bankrupt = newValue;
        }
        public void setAccountsReceivableTurnover(double newValue)
        {
            this.AccountsReceivableTurnover = newValue;
        }
        public void setOperatingGrossMargin(double newValue)
        {
            this.OperatingGrossMargin = newValue;
        }
        public void setCashCurrentLiability(double newValue)
        {
            this.CashCurrentLiability = newValue;
        }
    }
}
