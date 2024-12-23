using System.Collections.Generic;

namespace FileImport.Models
{
    internal class Reserve
    {
        internal string Ouz { get; set; }
        internal string Layer { get; set; }
        internal string Field { get; set; }
        internal string Year { get; set; }
        internal string MineralComponent { get; set; }
        internal List<Category> Categories { get; set; } = new List<Category>();

        // Добавленные поля залежей
        internal string DepositName { get; set; }
        internal string DepositType { get; set; }
        internal string DepositCollector { get; set; }
        internal string DepositMinDepth { get; set; }
        internal string DepositMaxDepth { get; set; }
        internal string DepositMinAbsDepth { get; set; }
        internal string DepositMaxAbsDepth { get; set; }
        internal string DepositYear { get; set; }
        internal string DepositProtocol { get; set; }
    }
}