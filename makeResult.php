<?php
require_once filter_input(INPUT_SERVER,'DOCUMENT_ROOT',FILTER_SANITIZE_STRING).'/general/config/config.php';
require_once LIBRARIES.'/PHPExcel/Classes/PHPExcel.php';
$objCRUD = new CRUD();

// Принимаем входные данные
$intFrom = strtotime($_POST['periodFrom'].' 00:00:00');
$intTo = strtotime($_POST['periodTo'].' 23:59:59');
$arrInputRes = array();
if (count($_POST['result'])){
    foreach ($_POST['result'] as $intKey => $intValue) $arrInputRes[] = $intValue*1;
}
$strRes = "'".implode("','",$arrInputRes)."'";
$intCnt = $_POST['cnt']*1;
if ($intCnt < 1){
    $intCnt = 1;
}

// Подготавливаем массив названий трансмиссий
$arrTransQuery = $objCRUD->customSelect("SELECT DISTINCT `trans_name`,`trans_1c_id` FROM `trans`");
$arrTrans = array();
foreach ($arrTransQuery as $arrTransRow){
    $arrTrans[$arrTransRow['trans_1c_id']] = $arrTransRow['trans_name'];
}
unset($arrTransQuery);

// Подготавливаем массив соответствия трансмиссий схемам
$arrTNamesQuery = $objCRUD->customSelect("SELECT DISTINCT `dform_trans_transid`,`dform_trans_transname` FROM `dform_trans`");
$arrTNames = array();
foreach ($arrTNamesQuery as $arrTNamesRow){
    $arrTNames[$arrTNamesRow['dform_trans_transname']][] = $arrTNamesRow['dform_trans_transid'];
}
foreach ($arrTNames as $strTName => $arrTNum){
    $arrTransmissions = array();
    foreach ($arrTNum as $intTransNum){
        if (!empty($arrTrans[$intTransNum])){
            $arrTransmissions[] = $arrTrans[$intTransNum];
        }
    }
    if (count($arrTransmissions)){
        sort($arrTransmissions);
        $arrTNames[$strTName] = implode(', ',$arrTransmissions);
    }
}
unset($arrTNamesQuery);
unset($arrTransmissions);

// Делаем выборку из stat_scheme
$arrPos = $objCRUD->customSelect("SELECT
        `stat_scheme_trans_1c_id` AS `trans`,
        `stat_scheme_scheme` AS `scheme`,
        `stat_scheme_pos` AS `pos`,
        `stat_scheme_userid` AS `userid`,
        `stat_scheme_result` AS `result`,
        MAX(`stat_scheme_time`) AS `timeLast`,
        COUNT(`stat_scheme_id`) AS `cnt`,
        `id`,
        `name`
    FROM `stat_scheme`
    LEFT JOIN `s_users` ON `s_users`.`s_users_id` = `stat_scheme`.`stat_scheme_userid`
    WHERE `stat_scheme_time` > {$intFrom}
      AND `stat_scheme_time` < {$intTo}
      AND `stat_scheme_result` IN ({$strRes})
    GROUP BY `scheme`,`pos`,`userid`");

if (count($arrPos)){

    // Определяем количество строк с одинаковыми scheme и pos
    $arrRowCnt = array();
    foreach ($arrPos as $arrRow){
        if (empty($arrRowCnt["{$arrRow['scheme']}"]["{$arrRow['pos']}"])){
            $arrRowCnt["{$arrRow['scheme']}"]["{$arrRow['pos']}"] = 1;
        }else{
            $arrRowCnt["{$arrRow['scheme']}"]["{$arrRow['pos']}"]++;
        }
    }
    foreach ($arrPos as $k => $arrRow){
        $arrPos["{$k}"]['rowsCnt'] = $arrRowCnt["{$arrRow['scheme']}"]["{$arrRow['pos']}"];
    }

    // Удаляем строки, количество которых меньше введённого в форме
    foreach ($arrPos as $k => $arrRow){
        if ($arrRow['rowsCnt'] < $intCnt){
            unset($arrPos[$k]);
        }
    }

}

if (count($arrPos)){ // Проверяем, не пустой ли массив после удаления строк
    // Заполняем столбец Трансмиссия из массива Трансмиссий, если он не пустой; или из массива соответствия трансмиссий схемам, если пустой
    foreach ($arrPos as $k => $arrRow){
        if (empty($arrRow['trans'])){
            $arrPos[$k]['trans'] = $arrTNames[$arrRow['scheme']];
        }else{
            $arrTransNum = explode(',',$arrRow['trans']);
            $arrTransmissions = array();
            foreach ($arrTransNum as $intTransNum){
                if (!empty($arrTrans[$intTransNum])){
                    $arrTransmissions[] = $arrTrans[$intTransNum];
                }
            }
            if (count($arrTransmissions)){
                sort($arrTransmissions);
                $arrPos[$k]['trans'] = implode(', ',$arrTransmissions);
            }
        }
    }

    // Определяем количество строк с одинаковыми scheme и pos и trans
    $arrTransRowsCnt = array();
    foreach ($arrPos as $arrRow){
        if (empty($arrTransRowsCnt["{$arrRow['scheme']}"]["{$arrRow['pos']}"]["{$arrRow['trans']}"])){
            $arrTransRowsCnt["{$arrRow['scheme']}"]["{$arrRow['pos']}"]["{$arrRow['trans']}"] = 1;
        }else{
            $arrTransRowsCnt["{$arrRow['scheme']}"]["{$arrRow['pos']}"]["{$arrRow['trans']}"]++;
        }
    }
    foreach ($arrPos as $k => $arrRow){
        $arrPos["{$k}"]['transRowsCnt'] = $arrTransRowsCnt["{$arrRow['scheme']}"]["{$arrRow['pos']}"]["{$arrRow['trans']}"];
    }

    // Сортируем массив по количеству одинаковых строк (scheme и pos), затем по количеству одинаковых trans внутри scheme и pos, затем по дате
    $arrRowsCnt = array();
    $arrSchemePosTransCnt = array();
    $arrScheme = array();
    $arrSchemePos = array();
    $arrTransName = array();
    $arrTimeLast = array();
    foreach ($arrPos as $k => $arrRow){
        $arrRowsCnt["{$k}"] = $arrRow['rowsCnt'];
        $arrSchemePosTransCnt["{$k}"] = $arrRow['transRowsCnt'];
        $arrScheme["{$k}"] = $arrRow['scheme'];
        $arrSchemePos["{$k}"] = $arrRow['pos'];
        $arrTransName["{$k}"] = $arrRow['trans'];
        $arrTimeLast["{$k}"] = $arrRow['timeLast'];
    }
    array_multisort($arrRowsCnt, SORT_DESC, $arrScheme, SORT_DESC, $arrSchemePos, SORT_ASC, $arrSchemePosTransCnt, SORT_DESC, $arrTransName, SORT_ASC, $arrTimeLast, SORT_DESC, $arrPos);
    unset($arrRowsCnt);
    unset($arrSchemePosTransCnt);
    unset($arrScheme);
    unset($arrSchemePos);
    unset($arrTimeLast);

    // Делаем экспорт в excel
    $objPHPExcel = new PHPExcel();

    $alignCenter = array(
        'alignment'=>array(
            'horizontal'=>PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            'vertical'=>PHPExcel_Style_Alignment::VERTICAL_TOP
        )
    );

    $objPHPExcel->getProperties()->setCreator("TRANSKIT")
        ->setLastModifiedBy("TRANSKIT")
        ->setTitle("Scheme positions")
        ->setSubject("Scheme positions")
        ->setDescription("Scheme positions")
        ->setKeywords("Scheme positions")
        ->setCategory("Scheme positions");

    $objPHPExcel->getDefaultStyle()->getFont()->setName('Arial');
    $objPHPExcel->getDefaultStyle()->getFont()->setSize(8);
    $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(30);
    $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(30);
    $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(30);
    $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(16);
    $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(17);
    $objPHPExcel->getActiveSheet()->freezePane('A2');

    $objPHPExcel->setActiveSheetIndex(0)
        ->setCellValue('A1', 'Трансмиссия')
        ->setCellValue('B1', 'Схема трансмиссии')
        ->setCellValue('C1', 'Позиция на схеме')
        ->setCellValue('D1', 'Имя клиента')
        ->setCellValue('E1', 'Код в 1С')
        ->setCellValue('F1', 'Внутр ID')
        ->setCellValue('G1', 'Кликов клиента')
        ->setCellValue('H1', 'Дата последнего клика клиента')
        ->setCellValue('I1', 'Соответствие');

    $intRow = 2;
    foreach ($arrPos as $k => $arrRow){
        $intRow++;
        if (($k > 0) and ($arrPos[$k-1]['pos'] != $arrRow['pos'])){
            $intRow++;
        }
        $strScheme = '<a href="'.MANAGEURL.'modules/schakpp/scheme.php?s='.$arrRow['scheme'].'">'.$arrRow['scheme'].'</a>';

        if ($arrRow['result'] == 1){
            $strRes = 'Деталь есть в базе';
        }else{
            $strRes = 'Детали нет в базе';
        }
        $dateTime = date('d.m.Y H:i',$arrRow['timeLast']);
        if (($arrRow['id']*1) < 1){
            $arrRow['id'] = 'нет';
        }
        $objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A'.$intRow, $arrRow['trans'])
            ->setCellValue('B'.$intRow, $arrRow['scheme'])
            ->setCellValue('C'.$intRow, $arrRow['pos'])
            ->setCellValue('D'.$intRow, $arrRow['name'])
            ->setCellValue('E'.$intRow, $arrRow['id'])
            ->setCellValue('F'.$intRow, $arrRow['userid'])
            ->setCellValue('G'.$intRow, $arrRow['cnt'])
            ->setCellValue('H'.$intRow, $dateTime)
            ->setCellValue('I'.$intRow, $strRes);

        $objPHPExcel->getActiveSheet()->getCell('A'.$intRow)->setValueExplicit($arrRow['trans'], PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->getCell('B'.$intRow)->setValueExplicit($arrRow['scheme'], PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->getCell('C'.$intRow)->setValueExplicit($arrRow['pos'], PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->getCell('D'.$intRow)->setValueExplicit($arrRow['name'], PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->getCell('E'.$intRow)->setValueExplicit($arrRow['id'], PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->getCell('F'.$intRow)->setValueExplicit($arrRow['userid'], PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->getCell('G'.$intRow)->setValueExplicit($arrRow['cnt'], PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->getCell('H'.$intRow)->setValueExplicit($dateTime, PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->getCell('I'.$intRow)->setValueExplicit($strRes, PHPExcel_Cell_DataType::TYPE_STRING);

        $objPHPExcel->setActiveSheetIndex(0)->getCell('B'.$intRow)->getHyperlink()->setUrl(MANAGEURL.'modules/schakpp/scheme.php?s='.$arrRow['scheme']);
        $objPHPExcel->setActiveSheetIndex(0)->getCell('C'.$intRow)->getHyperlink()->setUrl(MANAGEURL.'modules/schakpp/scheme.php?s='.$arrRow['scheme'].'&p='.$arrRow['pos']);
        $objPHPExcel->setActiveSheetIndex(0)->getCell('D'.$intRow)->getHyperlink()->setUrl(MANAGEURL.'redirect.php?m=admin&sm=susers&edit='.$arrRow['userid']);

        $objPHPExcel->setActiveSheetIndex(0)->getStyle('C'.$intRow)->applyFromArray($alignCenter);
        $objPHPExcel->setActiveSheetIndex(0)->getStyle('E'.$intRow)->applyFromArray($alignCenter);
        $objPHPExcel->setActiveSheetIndex(0)->getStyle('F'.$intRow)->applyFromArray($alignCenter);
        $objPHPExcel->setActiveSheetIndex(0)->getStyle('G'.$intRow)->applyFromArray($alignCenter);
    }

    $objPHPExcel->getActiveSheet()->setTitle('Клики позиций на схемах АКПП');
    $objPHPExcel->setActiveSheetIndex(0);

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
    $strFileName = 'schemePositionsClick-'.date('d.m.Y',time()).'.xls';
    $objWriter->save('xls/'.$strFileName);

    include('html/success.html');
}else{
    include('html/empty.html');
}
