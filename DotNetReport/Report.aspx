<%@ Page Title="" Language="C#" MasterPageFile="~/DotNetReport/ReportLayout.Master" AutoEventWireup="true" CodeBehind="Report.aspx.cs" Inherits="ReportBuilder.Demo.WebForms.DotNetReport.Report" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">    
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="scripts" runat="server">
    <script type="text/javascript">
    function printReport() {              
        var printWindow = window.open("");
        printWindow.document.open();
        printWindow.document.write('<html><head>'+
                                '<link href="/Content/bootstrap.css" rel="stylesheet" />'+
                                '<style type="text/css">a[href]:after {content: none !important;}</style>' +
                                '</head><body>' + $('.report-inner').html() +
                                '</body></html>');

        setTimeout(function(){
            printWindow.print();
            printWindow.close();
        }, 250);
    }

    function downloadExcel(currentSql, currentConnectKey, reportName) { 
        if (!currentSql) return;
        redirectToReport("/DotNetReport/ReportService.asmx/DownloadExcel", {
            reportSql: unescape(currentSql),
            connectKey: unescape(currentConnectKey),
            reportName: unescape(reportName)
        }, true, false);
    }

    $(document).ready(function () {
        var svc = "/DotNetReport/ReportService.asmx/";
        var vm = new reportViewModel({
            runReportUrl: svc+"Report",
            execReportUrl: svc+"RunReport",
            reportWizard: $("#filter-panel"),
            lookupListUrl: svc+"GetLookupList",
            apiUrl: svc + "CallReportApi",
            runReportApiUrl: svc +"RunReportApi",
            reportFilter: htmlDecode('<%= Model.ReportFilter %>'), 
            reportMode: "execute", 
            reportSql: "<%= Model.ReportSql %>", 
            reportConnect: "<%= Model.ConnectKey %>"
        });
        ko.applyBindings(vm);
        vm.LoadReport(<%= Model.ReportId %>, true);

        $(window).resize(function(){
            vm.DrawChart();            
        });
      });
    </script>
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="body" runat="server">
    
<div data-bind="with: ReportResult">
    
    <!-- ko ifnot: HasError -->
    <div class="report-view" data-bind="with: $root">
        <div class="pull-right">
            <a href="/DotNetReport/Index.aspx?folderId=<%= Model.SelectedFolder %>" class="btn btn-primary">
                Back to Reports
            </a>
            <a href="/DotNetReport/Index.aspx?reportId=<%= Model.ReportId %>&folderId=<%= Model.SelectedFolder%>" class="btn btn-primary">
                Edit Report
            </a>
            <button type="button" class="btn btn-default" onclick="printReport();">
                <span class="glyphicon glyphicon-print" aria-hidden="true"></span> Print Report
            </button>
            <a class="btn btn-default" data-bind="visible: !isChart() || ShowDataWithGraph(), click: downloadExcel(currentSql(), currentConnectKey(), ReportName())">
                <span class="fa fa-file-excel-o"></span> Export to Excel
            </a>

        </div>
        <br />
        <br />
        <div style="clear: both;"></div>
        <br />
        <div data-bind="template: { name: 'fly-filter-template' }"></div>
        <div data-bind="if: canDrilldown">
            <button class="btn btn-default btn-xs" data-bind="click: ExpandAll">Expand All</button>
            <button class="btn btn-default btn-xs" data-bind="click: CollapseAll">Collapse All</button>
            <br />
            <br />
        </div>
        <div class="report-menubar">
            <div class="col-xs-12 col-centered" data-bind="with: pager">
                <div class="form-inline" data-bind="visible: pages()">
                    <div class="form-group pull-left total-records">
                        <span data-bind="text: 'Total Records: ' + totalRecords()"></span><br />
                    </div>
                    <div class="pull-left">
                        <button class="btn btn-default btn-sm" onclick="downloadExcel();" data-bind="visible: !$root.isChart() || $root.ShowDataWithGraph()" title="Export to Excel">
                            <span class="fa fa-file-excel-o"></span> 
                        </button>
                    </div>
                    <div class="form-group pull-right">
                        <div data-bind="template: 'pager-template', data: $data"></div>
                    </div>
                </div>
            </div>
        </div> 
        <div class="report-canvas">   
            <div class="report-container">
                <div class="report-inner">
                    <h2 data-bind="text: ReportName"></h2>
                    <p data-bind="html: ReportDescription">
                    </p>
                    <div data-bind="with: ReportResult">
                        <div data-bind="template: 'report-template', data: $data"></div>
                    </div>
                </div>
            </div>            
        </div>
        <br />
        <span>Report ran on: <%=DateTime.Now.ToShortDateString() %> <%=@DateTime.Now.ToShortTimeString() %></span>         
    </div>
    <!-- /ko -->

    <!-- ko if: HasError -->
    <h2><%= Model.ReportName %></h2>
    <p>
        <%= Model.ReportDescription %>
    </p>

    <a href="/DotNetReport/Index.aspx?folderId=<%=Model.SelectedFolder %>" class="btn btn-primary">
        Back to Reports
    </a>
    <a href="/DotNetReport/Index.aspx?reportId=<%=Model.ReportId %>&folderId=<%=Model.SelectedFolder %>" class="btn btn-primary">
        Edit Report
    </a>
    <h3>An unexpected error occured while running the Report</h3>
    <hr />
    <b>Error Details</b>
    <p>
        <div data-bind="text: Exception"></div>
    </p>

    <!-- /ko -->

    <!-- ko if: ReportDebug || HasError -->    
        <br />
        <br />
        <hr />
        <code data-bind="text: ReportSql">
            
        </code>
    <!-- /ko -->
</div>


</asp:Content>