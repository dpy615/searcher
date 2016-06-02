using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using WebOfSciSearcher;

namespace SomallTest {
    class Program {
        static void Main(string[] args) {
            WebSearcher.SchoolSearchCommon common = new WebSearcher.SchoolSearchCommon();
            common.config = Utils.GetConfig()["Southampton"];
            common.GetData(1, "test.xls");
        }
    }
}
