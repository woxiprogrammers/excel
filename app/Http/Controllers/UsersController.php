<?php

namespace App\Http\Controllers;

use Carbon\Carbon;
use Illuminate\Http\Request;

use App\Http\Requests;
use App\Http\Controllers\Controller;

class UsersController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     * http://stackoverflow.com/questions/24740108/phpexcel-multiple-dropdown-list-that-dependent
     */
    public function index()
    {
        $objPHPExcel = new \PHPExcel();
        $objPHPExcel->getProperties()
            ->setCreator('Sagar Acharya')
            ->setTitle('PHPExcel Demo')
            ->setLastModifiedBy('Sagar Acharya')
            ->setDescription('A demo to show how to use PHPExcel to manipulate an Excel file')
            ->setSubject('PHP Excel manipulation')
            ->setKeywords('excel php office phpexcel lakers')
            ->setCategory('programming');

        $objPHPExcel->setActiveSheetIndex(0);
        $objPHPExcel->getActiveSheet('Worksheet 1')->SetCellValue("A1", "UK");
        $objPHPExcel->getActiveSheet('Worksheet 1')->SetCellValue("A2", "USA");

        $objPHPExcel->addNamedRange(
            new \PHPExcel_NamedRange(
                'countries',
                $objPHPExcel->getActiveSheet('Worksheet 1'),
                'A1:A2'
            )
        );

        $objPHPExcel->getActiveSheet('Worksheet 1')->SetCellValue("B1", "London");
        $objPHPExcel->getActiveSheet('Worksheet 1')->SetCellValue("B2", "Birmingham");
        $objPHPExcel->getActiveSheet('Worksheet 1')->SetCellValue("B3", "Leeds");
        $objPHPExcel->addNamedRange(
            new \PHPExcel_NamedRange(
                'UK',
                $objPHPExcel->getActiveSheet('Worksheet 1'),
                'B1:B3'
            )
        );

        $objPHPExcel->getActiveSheet('Worksheet 1')->SetCellValue("C1", "Atlanta");
        $objPHPExcel->getActiveSheet('Worksheet 1')->SetCellValue("C2", "New York");
        $objPHPExcel->getActiveSheet('Worksheet 1')->SetCellValue("C3", "Los Angeles");
        $objPHPExcel->addNamedRange(
            new \PHPExcel_NamedRange(
                'USA',
                $objPHPExcel->getActiveSheet('Worksheet 1'),
                'C1:C3'
            )
        );

        $objValidation = $objPHPExcel->getActiveSheet()->getCell('A1')->getDataValidation();
        $objValidation->setType( \PHPExcel_Cell_DataValidation::TYPE_LIST );
        $objValidation->setErrorStyle( \PHPExcel_Cell_DataValidation::STYLE_INFORMATION );
        $objValidation->setAllowBlank(false);
        $objValidation->setShowInputMessage(true);
        $objValidation->setShowErrorMessage(true);
        $objValidation->setShowDropDown(true);
        $objValidation->setErrorTitle('Input error');
        $objValidation->setError('Value is not in list.');
        $objValidation->setPromptTitle('Pick from list');
        $objValidation->setPrompt('Please pick a value from the drop-down list.');
        $objValidation->setFormula1("=countries"); //note this!


        $objValidation = $objPHPExcel->getActiveSheet()->getCell('B1')->getDataValidation();
        $objValidation->setType( \PHPExcel_Cell_DataValidation::TYPE_LIST );
        $objValidation->setErrorStyle( \PHPExcel_Cell_DataValidation::STYLE_INFORMATION );
        $objValidation->setAllowBlank(false);
        $objValidation->setShowInputMessage(true);
        $objValidation->setShowErrorMessage(true);
        $objValidation->setShowDropDown(true);
        $objValidation->setErrorTitle('Input error');
        $objValidation->setError('Value is not in list.');
        $objValidation->setPromptTitle('Pick from list');
        $objValidation->setPrompt('Please pick a value from the drop-down list.');
        $objValidation->setFormula1('=INDIRECT($A$1)');
        $path = $_SERVER['DOCUMENT_ROOT'].'/uploads/';
        $objWriter = new \PHPExcel_Writer_Excel2007($objPHPExcel);
        $name = strtotime(Carbon::now());
        $name = $name.'.xlsx';
        $objWriter->save($path.$name);
        chmod($path.$name,0777);
        //dd($objPHPExcel);
    }

    /**
     * Show the form for creating a new resource.
     *
     * @return \Illuminate\Http\Response
     * https://docs.typo3.org/typo3cms/extensions/phpexcel_library/1.7.4/manual.html
     */
    public function create()
    {
        $data = implode (", ", array('UK,USA'));
        $objPHPExcel = new \PHPExcel();
        $objPHPExcel->getProperties()
            ->setCreator('Sagar Acharya')
            ->setTitle('PHPExcel Demo')
            ->setLastModifiedBy('Sagar Acharya')
            ->setDescription('A demo to show how to use PHPExcel to manipulate an Excel file')
            ->setSubject('PHP Excel manipulation')
            ->setKeywords('excel php office phpexcel lakers')
            ->setCategory('programming');

        $objPHPExcel->setActiveSheetIndex(0);
        $objPHPExcel->addNamedRange(
            new \PHPExcel_NamedRange(
                'UK',
                $objPHPExcel->getActiveSheet('Worksheet 1'),
                'Leeds'
            )
        );
        $objPHPExcel->addNamedRange(
            new \PHPExcel_NamedRange(
                'USA',
                $objPHPExcel->getActiveSheet('Worksheet 1'),
                'LA'
            )
        );
        $cities = implode (":", array('London,Leeds,DC,LA'));
        $objValidation = $objPHPExcel->getActiveSheet()->getCell('A1')->getDataValidation();
        $objValidation->setType( \PHPExcel_Cell_DataValidation::TYPE_LIST );
        $objValidation->setErrorStyle( \PHPExcel_Cell_DataValidation::STYLE_INFORMATION );
        $objValidation->setAllowBlank(false);
        $objValidation->setShowInputMessage(true);
        $objValidation->setShowErrorMessage(true);
        $objValidation->setShowDropDown(true);
        $objValidation->setErrorTitle('Input error');
        $objValidation->setError('Value is not in list.');
        $objValidation->setPromptTitle('Pick from list');
        $objValidation->setPrompt('Please pick a value from the drop-down list.');
        $objValidation->setFormula1('"'.$data.'"');
        $objPHPExcel->getActiveSheet()->getCell('A1')->setDataValidation($objValidation);

        /////////////////////////
        $objValidation = $objPHPExcel->getActiveSheet()->getCell('B1')->getDataValidation();
        $objValidation->setType( \PHPExcel_Cell_DataValidation::TYPE_LIST );
        $objValidation->setErrorStyle( \PHPExcel_Cell_DataValidation::STYLE_INFORMATION );
        $objValidation->setAllowBlank(false);
        $objValidation->setShowInputMessage(true);
        $objValidation->setShowErrorMessage(true);
        $objValidation->setShowDropDown(true);
        $objValidation->setErrorTitle('Input error');
        $objValidation->setError('Value is not in list.');
        $objValidation->setPromptTitle('Pick from list');
        $objValidation->setPrompt('Please pick a value from the drop-down list.');
        $objValidation->setFormula1('=INDIRECT($A$1)');
        $objPHPExcel->getActiveSheet()->getCell('B1')->setDataValidation($objValidation);

        $objWriter = new \PHPExcel_Writer_Excel2007($objPHPExcel);
        $path = $_SERVER['DOCUMENT_ROOT'].'/uploads/';
        $objWriter->save($path.'some_excel_file.xlsx');
    }

    /**
     * Store a newly created resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @return \Illuminate\Http\Response
     */
    public function store(Request $request)
    {
        //
    }

    /**
     * Display the specified resource.
     *
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function show($id)
    {
        //
    }

    /**
     * Show the form for editing the specified resource.
     *
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function edit($id)
    {
        //
    }

    /**
     * Update the specified resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function update(Request $request, $id)
    {
        //
    }

    /**
     * Remove the specified resource from storage.
     *
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function destroy($id)
    {
        //
    }
}
