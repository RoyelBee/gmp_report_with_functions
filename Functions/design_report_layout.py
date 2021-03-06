import Functions.item_wise_yesterday_sales as yesterday
import Functions.no_sales_record as noSales
import Functions.no_stock_record as noStock
import Functions.read_gpm_info as gpm
import Functions.sales_and_stock_record as SalesStock

import Functions.branch_wise_stocks as branch_stock
import Functions.branch_stock_summery as bs


def generate_layout(gpm_name):
    # print('GPM Name  = ', gpm_name)
    results = """ <!DOCTYPE html>
            <html>
            <head>
                <style>
                    table {
                        font-family: arial, sans-serif;
                        border-collapse: collapse;
                        padding: 0px;
                        white-space: nowrap;
                        font-size: 10px;
                        border: 1px solid black;
                    }
                    th {
                        color: black;
                        border: 1px solid black;
                        font-size: 10px;
                        padding-left: 5px;
                        padding-right: 5px;
                    }
                    .central{
                    text-align: center;
                    }
                    th.style1{
                        background-color: #e5f0e5;
                        padding-left: 2px;
                        padding-right: 5px;
                        text-align: right;
                    }
                    .style1{
                        padding-right: 3px;
                        text-align: right;
                    }
                    .item_sl{
                        background-color: #e5f0e5;
                        padding-right: 3px;
                        text-align: center;
                    }
                    td.serial{
                        padding-left: 5px;
                        text-align: left;
                    }
                    .brand {
                        background-color: #e5f0e5;
                        font-weight: bold;
                        color: black;
                        text-align: left;
                    }
                    .brandtd {
                        text-align: left;
                        padding-left: 2px;
                        font-weight: bold;
                        font-size: 12px;
                        color: black;
                    }
                    th.style2{
                        background-color: #faeaca;
                    }
                    .remaining{
                        background-color: #faeaca;
                        padding-left: 20px;
                    }
                    .sales_monthly_trend{
                     background-color: #faeaca;
                     padding-left: 20px;
    
                    }
                    .style3{
                        background-color: #f7e0b3;
                        padding-left: 5px;
                        padding-right: 2px;
    
                    }
                    td {
                        font-family: "Tohoma";
                        border: 1px solid gray;
                    }
                    .branch_bg_color{
                        font-family: "Tohoma";
                        background-color: #a7c496;
                        padding-left: 5px;
                        color: black;
                        padding-right: 2px;
                        text-align: right;
                        font-size: 8px;
    
                    }
                    td.num_style {
                        text-align: right;
                        border: 1px solid gray;
                        padding-right: 5px;
                    }
                    .color_style{
                        padding-left: 10px;
                        padding-top: 5px;
                        padding-bottom: 5px;
                        padding-right: 10px;
                        color: black;
                        width: 25%
                    }
                    .color_style{
                        width: 16.66%;
                        padding-left: 10px;
                        padding-top: 5px;
                        padding-bottom: 10px;
                        padding-right: 10px;
                        font-size: 14px;
                        font-weight: bolder;
                    }
                      .my_rotate {
                            color: red !important;
                      }
                    .nation_wide{
                        background-color: #ddd9d9;
                        text-align: right;
                        padding-right: 2px;
                        border: 1px solid gray;
                        padding-left: 10px;
                        font-weight: bolder;
                        font-size: 12px;
                        color: black;
                    }
                    th.description{
                        width: 900px !important;
                        background-color: #e5f0e5;
                        font-size: 12px;
                        color: black;
                    }
                    .description1{
                        width: 250px !important;
                        background-color: #e5f0e5;
                        font-size: 12px;
                        color: black;
                    }
                    .descriptiontd{
                    padding-left: 2px;
                    }
                    .uom{
                        background-color: #e5f0e5;
                        font-size: 11px;
                        color: black;
                    }
    
                    .my_margin{
                    margin-right: 220px;
                    color: #E5F0E5 !important;
                    }
                    th.avg_sales{
                        width: 500px !important;
                        background-color: #e5f0e5;
                        font-size: 12px;
                        color: black;
                    }
                     td.avg_sales{
                        width: 100% !important;
                        padding-left:5px;
                    }
                    .remarks{
                    padding-left: 1px;
                    padding-right: 1px;
                    text-align: right;
                    }
                    .info{
                        background-color: #a9f2e7;
                        font-size: 13px;
                        text-align: left;
                        text-decoration-line: none;
                        color: black;
    
                    }
    
                    .banner{
                    width: 800px !important;
                    height: 200px !important;
                    border: 1px solid #0f6674;
                    }
                    
                .float_left {
                    width: 50%;
                    float: left;
                }
    
                </style>
    
            </head>
            <body>
                <img src="cid:banner_ai"> <br>
                <img src="cid:dash"> <br>
                <img src="cid:cm"> <br>
                <img src="cid:executive"> <br>
                <img src="cid:brand"> <br> <br>
                
                <table border="1px solid gray" width="79%">
                    <tr>
                        <th colspan="6" style="background-color: #0beb9b "><h1>Item wise Yesterday Sales 
                        Quantity </h1> </th>
                    </tr>
                    
                    <tr>
                        <th rowspan="2" class="brand">BSL<br> No.</th>
                        <th rowspan="2" class="brand"> Brand &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</th>
                        <th rowspan="2" class="item_sl">Item SL</th>
                        <th rowspan="2" class="description1">Item Description</th>
                        <th rowspan="2" class="uom" style="text-align: right"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; UOM</th>
                        <th rowspan="2" class="uom" style="text-align: right"> Yesterday Sales</th>
                    </tr>
                
                   <tr> """ + yesterday.item_wise_yesterday_sales_Records()+ """
                </table>  <br> <br>
            
            <table border="1px solid gray" width="79%">
                    <tr>
                        <th colspan="5" style="background-color: #e3f865 "><h1> Yesterday No Sales Item </h1> </th>
                    </tr>
                    
                    <tr>
                        <th rowspan="2" class="brand">BSL<br> No.</th>
                        <th rowspan="2" class="brand"> Brand &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</th>
                        <th rowspan="2" class="item_sl">Item SL</th>
                        <th rowspan="2" class="description1">Item Description</th>
                        <th rowspan="2" class="uom" style="text-align: right"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; UOM</th>
                        
                    </tr>
                
                   <tr> """ + yesterday.item_wise_yesterday_no_sales_Records()+ """
                </table>  <br> <br>

                
                <table border="1px solid gray" width="79%">
                    <tr>
                        <th colspan="5" style="background-color: #cbe14c" > <h1>No Sales Item: Last 3 Months</h1></th>
                    </tr>
                    <tr>
                        <th rowspan="2" class="brand">BSL No</th>
                        <th rowspan="2" class="brand"> Brand &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</th>
                        <th rowspan="2" class="item_sl">Item SL</th>
                        <th rowspan="2" class="description1">Item Description</th>
                        <th rowspan="2" class="uom" style="text-align: right"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; UOM</th>
                    </tr>
    
                    <tr> """ + noSales.get_No_Sales_Records() + """
                    </table> 
           
           <br> <br>
                <table border="1px solid gray" width="79%">
                    <tr>
                        <th colspan="7" style="background-color: #faac9f;"><h1> No Stocks Item: Last 3 Months</h1></th>
                    </tr>
                    <tr>
                        <th rowspan="2" class="brand">BSL No</th>
                        <th rowspan="2" class="brand"> Brand &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</th>
                        <th rowspan="2" class="item_sl">Item SL</th>
                        <th rowspan="2" class="description1">Item Description</th>
                        <th rowspan="2" class="uom" style="text-align: right"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; UOM</th>
                        <th rowspan="2" class="uom">Total Ordered</th>
                        <th rowspan="2" class="uom" style="text-align: right"> Estimated Sales</th>
                    </tr>
                    <tr> """ + noStock.get_No_Stock_Records() + """  </tr>
                </table>  <br> <br>
        
        <table border="1px solid gray" width="79%">
        <tr>
            <th colspan="7" style="background-color: #cbe14c"><h1>Branch Wise Item Stock Category</h1></th>
        </tr>
        <tr>
            <th rowspan="2" class="brand">Branch</th>
            <th rowspan="2" class="brand"> Nill</th>
            <th rowspan="2" class="item_sl">Super Under Stock</th>
            <th rowspan="2" class="description1">Under Stock</th>
            <th rowspan="2" class="uom" style="text-align: right"> Normal Stock</th>
            <th rowspan="2" class="uom" style="text-align: right">Over Stock</th>
            <th rowspan="2" class="uom" style="text-align: right">Super Over Stock</th>
        </tr>

        <tr> """ + bs.branch_wise_nil_us_ss() + """ 
        
        </table>
        <br> <br>


            <table border="1px solid gray" width="77%">
                <tr>
                    <th colspan="5" style="background-color: #34ce57;"><h1> Branch Wise Stock</h1></th>
                    <th colspan="5" style="background-color: #ff2300" class="color_style"> Nill</th>
                    <th colspan="5" style="background-color: #ff971a" class="color_style">Super Under Stock</th>
                    <th colspan="5" style="background-color: #eee298;" class="color_style">Under Stock</th>
                    <th colspan="5" class="color_style">Normal Stock</th>
                    <th colspan="5" style="background-color: #cbe14c; color: black"   class="color_style">Over 
                    Stock</th>
                    <th colspan="6" style="background-color: #fff900; color: black"  class="color_style">Super Over 
                    Stock</th>
                </tr>
                
                <tr>
                    <th rowspan="2" class="brand">BSL<br> No.</th>
                    <th rowspan="2" class="brand"> Brand &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</th>
                    <th rowspan="2" class="item_sl">Item SL</th>
                    <th rowspan="2" class="description"> <div class="my_margin">.</div> Item Description</th>
                    <th rowspan="2" class="uom" style="text-align: right"> &nbsp;&nbsp;&nbsp;&nbsp;UOM</th>
                    <th rowspan="2" class="uom">BOG</th>
                    <th rowspan="2" class="uom">BSL</th>
                    <th rowspan="2" class="uom">COM</th>
                    <th rowspan="2" class="uom">COX</th>
                    <th rowspan="2" class="uom">CTG</th>
                    <th rowspan="2" class="uom">CTN</th>
                    <th rowspan="2" class="uom">DNJ</th>
                    <th rowspan="2" class="uom">FEN</th>
                    <th rowspan="2" class="uom">FRD</th>
                    <th rowspan="2" class="uom">GZP</th>
                    <th rowspan="2" class="uom">HZJ</th>
                    <th rowspan="2" class="uom">JES</th>
                    <th rowspan="2" class="uom">KHL</th>
                    <th rowspan="2" class="uom">KRN</th>
                    <th rowspan="2" class="uom">KSG</th>
                    <th rowspan="2" class="uom">KUS</th>
                    <th rowspan="2" class="uom">MHK</th>
                    <th rowspan="2" class="uom">MIR</th>
                    <th rowspan="2" class="uom">MLV</th>
                    <th rowspan="2" class="uom">MOT</th>
                    <th rowspan="2" class="uom">MYM</th>
                    <th rowspan="2" class="uom">NAJ</th>
                    <th rowspan="2" class="uom">NOK</th>
                    <th rowspan="2" class="uom">PAT</th>
                    <th rowspan="2" class="uom">PBN</th>
                    <th rowspan="2" class="uom">RAJ</th>
                    <th rowspan="2" class="uom">RNG</th>
                    <th rowspan="2" class="uom">SAV</th>
                    <th rowspan="2" class="uom">SYL</th>
                    <th rowspan="2" class="uom">TGL</th>
                    <th rowspan="2" class="uom">VRB</th>     
                </tr>
                <tr> """ + branch_stock.branch_wise_stocks_Records() + """
            </table>  <br> <br>


           
                <table border="1px solid gray" cellspacing ="20">
                 <tr>
                    <th colspan="15" class="info" style="text-align: center"> """ + gpm.getGPMNFullInfo(gpm_name) + """
                    </th>
                    <th colspan="3" style="font-weight: bolder; font-size: 12px; background-color: #e6a454 ">SKF Plant</th>
                    <th rowspan="3" style="background-color: #d0ff89"><div>TDCL Central WH</div></th>
                    <th rowspan="3" style="background-color: #95ff89"><div>Branch Total</div></th>
                    <th colspan="5" style="background-color: #ff2300" class="color_style"> Nill</th>
                    <th colspan="5" style="background-color: #ff971a" class="color_style">Super Under Stock</th>
                    <th colspan="5" style="background-color: #eee298;" class="color_style">Under Stock</th>
                    <th colspan="5" class="color_style">Normal Stock</th>
                    <th colspan="5" style="background-color: #cbe14c; color: black"   class="color_style">Over Stock</th>
                    <th colspan="6" style="background-color: #fff900; color: black"  class="color_style">Super Over Stock</th>
                </tr>
    
                <tr>
                    <th rowspan="2" class="style1">BSL<br> No.</th>
                    <th rowspan="2" class="brand"> Brand &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</th>
                    <th rowspan="2" class="item_sl">Item SL</th>
                    <th rowspan="2" class="description"> <div class="my_margin">.</div> Item Description</th>
                    <th rowspan="2" class="uom"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; UOM &nbsp;</th>
                    <th rowspan="2" class="sales_monthly_trend">Yesterday Sales </th>
                    <th rowspan="2" class="sales_monthly_trend">Avg Sales Per Day </th>
                    <th rowspan="2" class="sales_monthly_trend">Monthly Sales Target</th>
                    <th rowspan="2" class="style2" ><div>MTD Sales Target</div></th>
                    <th rowspan="2" class="style2"><div>Actual Sales MTD</div></th>
                    <th rowspan="2" class="style2"><div>MTD Sales Achv %</div></th>
                    <th rowspan="2" class="style2"><div>Monthly Sales Achv %</div></th>
                    <th rowspan="2" class="style2"><div>Monthly Sales Trend</div></th>
                    <th rowspan="2" class="style2"> <div>Monthly Sales Trend Achv</div></th>
                    <th rowspan="2" class="remaining"><div>Remaining Quantity</div></th>
                    <th rowspan="2" class="nation_wide"><div>Nationwide Stock</div></th>
                    <th class="style3" rowspan="2">SKF <br>Mirpur </th>
                    <th class="style3" rowspan="2">SKF <br>Rupganje </th>
                    <th class="style3" rowspan="2">SKF <br>Tongi </th>
                    <th colspan="31" style="font-weight: bolder; font-size: 14px"> TDCL Branches</th>
                </tr>
                <tr>
    
                    <th class="branch_bg_color">BOG</th>
                    <th class="branch_bg_color">BSL</th>
                    <th class="branch_bg_color">COM</th>
                    <th class="branch_bg_color">COX</th>
                    <th class="branch_bg_color">CTG</th>
                    <th class="branch_bg_color">CTN</th>
                    <th class="branch_bg_color"> DNJ</th>
                    <th class="branch_bg_color">FEN</th>
                    <th class="branch_bg_color">FRD</th>
                    <th class="branch_bg_color">GZP</th>
                    <th class="branch_bg_color">HZJ</th>
                    <th class="branch_bg_color">JES</th>
                    <th class="branch_bg_color">KHL</th>
                    <th class="branch_bg_color">KRN</th>
                    <th class="branch_bg_color">KSG</th>
                    <th class="branch_bg_color">KUS</th>
                    <th class="branch_bg_color">MHK</th>
                    <th class="branch_bg_color">MIR</th>
                    <th class="branch_bg_color">MLV</th>
                    <th class="branch_bg_color" style="font-size: 8px;"> MOT</th>
                    <th class="branch_bg_color">MYM</th>
                    <th class="branch_bg_color">NAJ</th>
                    <th class="branch_bg_color">NOK</th>
                    <th class="branch_bg_color">PAT</th>
                    <th class="branch_bg_color">PBN</th>
                    <th class="branch_bg_color">RAJ</th>
                    <th class="branch_bg_color" style="font-size: 8px;">RNG</th>
                    <th class="branch_bg_color">SAV</th>
                    <th class="branch_bg_color">SYL</th>
                    <th class="branch_bg_color">TGL</th>
                    <th class="branch_bg_color">VRB</th>
    
                    </tr>
    
                    <tr>""" + SalesStock.get_Sales_and_Stock_Records() + """
                </table> <br>
                                
                </body>
            </html>
        """
    return results

# """ + gpm.getGPMNFullInfo(gpm_name) + """