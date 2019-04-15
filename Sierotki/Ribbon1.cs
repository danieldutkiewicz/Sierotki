using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace Sierotki
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            button.ScreenTip = "Napraw wiersze";
            button.SuperTip = "Jednoliterowe spójniki i przyimki jak: w, i, u, o, a, z, należy przesunąć na początek wiersza. \n \n" +
                "Pozostawienie ich na końcu wiersza jest błędem stylistycznym, którego należy unikać. \n \n" +
                "Powyższy dodatek przesuwa wszystkie samotne litery na początek kolejnego wiersza jednym kliknięciem.";
        }

        private void Sierotki_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.SearchReplace();
        }
    }
}
