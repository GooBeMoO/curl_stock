<?php
//header("Content-Type:text/html; charset=utf-8");

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

//初始化
$ch = curl_init();

//設置選項
$useagent="Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36";

//轉為unix time 13位制
$time=microtime(true).'<br>';
$time=str_replace(".","",$time);
$unixtime=substr($time,0,13);
//echo $unixtime;

//設定抓取URL
curl_setopt($ch, CURLOPT_URL, "http://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=tse_2313.tw&json=1&delay=0&_=$unixtime");

//抓取結果直接返回(為0,直接輸出內容)
curl_setopt($ch, CURLOPT_RETURNTRANSFER,1);

curl_setopt($ch, CURLOPT_ENCODING, 'gzip');

curl_setopt($ch,CURLOPT_USERAGENT,$useagent);

 $output = curl_exec($ch);

 //json轉為陣列
 $outputarray = json_decode($output,true);



$sellprice=$outputarray['msgArray'][0]['a'];
$buyprice=$outputarray['msgArray'][0]['b'];
$sellnumber=$outputarray['msgArray'][0]['f'];
$buynumber=$outputarray['msgArray'][0]['g'];
$company_name=$outputarray['msgArray'][0]['nf'];
$company_symbol=$outputarray['msgArray'][0]['c'];

$tostr_spilt2=array($sellprice,$buyprice,$sellnumber,$buynumber);
$arr2=count($tostr_spilt2);

/*for($data=0;$data<$arr2;$data++){
   $result=preg_split("/\_/",$tostr_spilt2[$data]);
   $tostr_spilt2[$data]=$result;
}
*/
print_r($tostr_spilt2);


$spreadsheet = new Spreadsheet();
// 寫入文檔

$arrayTitle=   [ '公司名稱'    ,'代碼'         ,'賣出價格'         ,'買入價格'     ,'賣出數量'       ,'買入數量'        ]  ;

$spreadsheet->getActiveSheet()->fromArray($arrayTitle,NULL,'A1');

$arrayData =   [ $company_name,$company_symbol, $tostr_spilt2[0],$tostr_spilt2[1],$tostr_spilt2[2],$tostr_spilt2[3] ]   ;

$spreadsheet->getActiveSheet()->fromArray($arrayData,NULL,'A2');

//
$arrayrowsize=['A','C','D','E','F'];
for($a=0;$a<count($arrayrowsize);$a++){
$spreadsheet->getActiveSheet()->getColumnDimension($arrayrowsize[$a])->setWidth(30);
}
$spreadsheet->getActiveSheet()->getStyle('B2')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
// 儲存文檔
$writer = new Xlsx($spreadsheet);
$writer->save('stock.xlsx');


if ($output === FALSE) {
    echo "cURL Error: " . curl_error($ch);
}

$info = curl_getinfo($ch);

//echo '獲取'.$info['url'];

curl_close($ch);



