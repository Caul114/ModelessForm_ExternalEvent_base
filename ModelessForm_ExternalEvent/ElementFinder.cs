using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

public class ElementFinder
{
    public static List<Element> FindAllElements(Document doc)
    {
        List<Element> list = new List<Element>();

        FilteredElementCollector finalCollector = new FilteredElementCollector(doc);

        ElementIsElementTypeFilter filter1 = new ElementIsElementTypeFilter(false);
        finalCollector.WherePasses(filter1);
        ElementIsElementTypeFilter filter2 = new ElementIsElementTypeFilter(true);
        finalCollector.UnionWith((new FilteredElementCollector(doc)).WherePasses(filter2));

        list = finalCollector.ToList<Element>();

        return list;
    }
}
