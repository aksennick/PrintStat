using System.Web;
using System.Web.Optimization;

namespace PrintStat
{
    public class BundleConfig
    {
        // Дополнительные сведения о Bundling см. по адресу http://go.microsoft.com/fwlink/?LinkId=254725
        public static void RegisterBundles(BundleCollection bundles)
        {
            bundles.Add(new ScriptBundle("~/bundles/jquery").Include(
                        "~/Scripts/jquery-{version}.js"));


            bundles.Add(new ScriptBundle("~/bundles/jqueryui").Include(
                        "~/Scripts/jquery-ui-1.*"));

            bundles.Add(new ScriptBundle("~/bundles/jqueryval").Include(
                        "~/Scripts/jquery.unobtrusive*",
                        "~/Scripts/jquery.validate*"));
            
            // Используйте версию Modernizr для разработчиков, чтобы учиться работать. Когда вы будете готовы перейти к работе,
            // используйте средство построения на сайте http://modernizr.com, чтобы выбрать только нужные тесты.
            bundles.Add(new ScriptBundle("~/bundles/modernizr").Include(
                        "~/Scripts/modernizr-*"));

            bundles.Add(new StyleBundle("~/Content/css").Include("~/Content/site.css"));

            bundles.Add(new StyleBundle("~/Content/themes/base/images/css").Include(
                        "~/Content/themes/base/images/core.css",
                        "~/Content/themes/base/images/resizable.css",
                        "~/Content/themes/base/images/selectable.css",
                        "~/Content/themes/base/images/accordion.css",
                        "~/Content/themes/base/images/autocomplete.css",
                        "~/Content/themes/base/images/button.css",
                        "~/Content/themes/base/images/dialog.css",
                        "~/Content/themes/base/images/slider.css",
                        "~/Content/themes/base/images/tabs.css",
                        "~/Content/themes/base/images/datepicker.css",
                        "~/Content/themes/base/images/progressbar.css",
                        "~/Content/themes/base/images/theme.css",
                        "~/Content/themes/base/images/all.css",
                        "~/Content/themes/base/images/base.css",
                        "~/Content/themes/base/images/draggable.css",
                        "~/Content/themes/base/images/menu.css",
                        "~/Content/themes/base/images/selectmenu.css",
                        "~/Content/themes/base/images/sortable.css",
                        "~/Content/themes/base/images/spinner.css",
                        "~/Content/themes/base/images/tooltip.css"
                       
                        ));
            bundles.Add(new StyleBundle("~/Content/DialogCSS").Include(
                        "~/Content/themes/jquery-ui.css"
                ));

            bundles.Add(new StyleBundle("~/Content/datetimepickerCSS").Include(
            "~/Content/themes/jquery.datetimepicker.css"
            ));

            bundles.Add(new StyleBundle("~/Content/multiselectCSS").Include(
                "~/Content/themes/jquery.multi-select.css"));


            bundles.Add(new ScriptBundle("~/bundles/Dialog").Include(
                        "~/Script/openDialog.js" ));

            bundles.Add(new ScriptBundle("~/bundles/datetimepicker").Include(
                        "~/Script/jquery.datetimepicker.js" ));

            bundles.Add(new ScriptBundle("~/bundles/multiselect").Include(
                "~/Scripts/jquery.multi-select.js"));

            bundles.Add(new ScriptBundle("~/bundles/quicksearch").Include(
                "~/Scripts/jquery.quicksearch.js"));
          
        }
    }
}