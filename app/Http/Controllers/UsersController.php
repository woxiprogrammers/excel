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
     * $objWriter->save('php://output');
     * https://gist.github.com/r-sal/4313500
     * http://www.tisuchi.com/use-phpexcel-library/
     */
    public function index()
    {
        $objPHPExcel = new \PHPExcel();
        $objPHPExcel->createSheet(1);
        $objPHPExcel->getProperties()
            ->setCreator('Sagar Acharya')
            ->setTitle('PHPExcel Demo')
            ->setLastModifiedBy('Sagar Acharya')
            ->setDescription('A demo to show how to use PHPExcel to manipulate an Excel file')
            ->setSubject('PHP Excel manipulation')
            ->setKeywords('excel php office phpexcel lakers')
            ->setCategory('programming');

        $objPHPExcel->setActiveSheetIndex(1);


        /*$objPHPExcel->getActiveSheet('Worksheet 1')->SetCellValue("A1", "UK");
        $objPHPExcel->getActiveSheet('Worksheet 1')->SetCellValue("A2", "USA");

        $objPHPExcel->addNamedRange(
            new \PHPExcel_NamedRange(
                'countries',
                $objPHPExcel->getActiveSheet('Worksheet 2'),
                'A1:A2'
            )
        );*/
        $data = implode (", ", array('UK,USA'));

        $objPHPExcel->getActiveSheet('Worksheet 2')->SetCellValue("B1", "London");
        $objPHPExcel->getActiveSheet('Worksheet 2')->SetCellValue("B2", "Birmingham");
        $objPHPExcel->getActiveSheet('Worksheet 2')->SetCellValue("B3", "Leeds");
        $objPHPExcel->addNamedRange(
            new \PHPExcel_NamedRange(
                'UK',
                $objPHPExcel->getActiveSheet('Worksheet 2'),
                'B1:B3'
            )
        );

        $objPHPExcel->getActiveSheet('Worksheet 2')->SetCellValue("C1", "Atlanta");
        $objPHPExcel->getActiveSheet('Worksheet 2')->SetCellValue("C2", "New York");
        $objPHPExcel->getActiveSheet('Worksheet 2')->SetCellValue("C3", "Los Angeles");
        $objPHPExcel->addNamedRange(
            new \PHPExcel_NamedRange(
                'USA',
                $objPHPExcel->getActiveSheet('Worksheet 2'),
                'C1:C3'
            )
        );
        $objPHPExcel->setActiveSheetIndex(0);
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
        $objValidation->setFormula1(".$data."); //note this!


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
     * http://stackoverflow.com/questions/18377976/read-the-list-of-options-in-a-drop-down-list-phpexcel
     * fully functional code for multiple dropdown list
     */
    public function show()
    {
        $objPHPExcel = new \PHPExcel();
        $newSheet=$objPHPExcel->createSheet();
        $objPHPExcel->setActiveSheetIndex(1);
        $newSheet->setTitle("CountriesList");

        $objPHPExcel->setActiveSheetIndex(1)
            ->SetCellValue("A1", "UK")
            ->SetCellValue("A2", "USA");
            /*->SetCellValue("A3", "CANADA")
            ->SetCellValue("A4", "INDIA")
            ->SetCellValue("A5", "POLAND")
            ->SetCellValue("A6", "ENGLAND");// Drop down data in sheet 1*/
        $objPHPExcel->setActiveSheetIndex(1)
            ->SetCellValue("B1", "London")
            ->SetCellValue("B2", "Birmingham")
            ->SetCellValue("B3", "Leeds");
        $objPHPExcel->setActiveSheetIndex(1)
            ->SetCellValue("C1", "LA")
            ->SetCellValue("C2", "NY");
        $objPHPExcel->addNamedRange(
            new \PHPExcel_NamedRange(
                'countries',
                $objPHPExcel->setActiveSheetIndex(1),
                'A1:A6'
            )
        );
        $objPHPExcel->addNamedRange(
            new \PHPExcel_NamedRange(
                'UK',
                $objPHPExcel->setActiveSheetIndex(1),
                'B1:B3'
            )
        );
        $objPHPExcel->addNamedRange(
            new \PHPExcel_NamedRange(
                'USA',
                $objPHPExcel->setActiveSheetIndex(1),
                'C1:C2'
            )
        );

        /* Lock Cells */

        $objPHPExcel->getActiveSheet()->getProtection()->setSheet(true);    // Needs to be set to true in order to

        $objPHPExcel->getActiveSheet()
            ->getStyle('A1:A2', 'B1:B3','C1:C2')
            ->getProtection()
            ->setLocked(
                \PHPExcel_Style_Protection::PROTECTION_PROTECTED
            );
        
        /* End Lock */


        $objPHPExcel->setActiveSheetIndex(0)->SetCellValue("A1", "UK");

        $objPHPExcel->setActiveSheetIndex(0);// Drop down in sheet 0
        $objValidation = $objPHPExcel->getSheet(0)->getCell('A1')->getDataValidation();
        $objValidation->setType( \PHPExcel_Cell_DataValidation::TYPE_LIST );
        $objValidation->setErrorStyle( \PHPExcel_Cell_DataValidation::STYLE_INFORMATION );
        $objValidation->setAllowBlank(false);
        $objValidation->setShowInputMessage(true);
        $objValidation->setShowErrorMessage(true);
        $objValidation->setShowDropDown(true);
        $objValidation->setErrorTitle('Input error');
        $objValidation->setError('Value is not in list.');
        $objValidation->setFormula1("=countries");
        $objPHPExcel->getActiveSheet()->getCell('A1')->setDataValidation($objValidation);


        //$objPHPExcel->setActiveSheetIndex(0)->SetCellValue("B1", "London");

        $objPHPExcel->setActiveSheetIndex(0);// Drop down in sheet 0
        $objValidation = $objPHPExcel->getSheet(0)->getCell('B1')->getDataValidation();
        $objValidation->setType( \PHPExcel_Cell_DataValidation::TYPE_LIST );
        $objValidation->setErrorStyle( \PHPExcel_Cell_DataValidation::STYLE_INFORMATION );
        $objValidation->setAllowBlank(false);
        $objValidation->setShowInputMessage(true);
        $objValidation->setShowErrorMessage(true);
        $objValidation->setShowDropDown(true);
        $objValidation->setErrorTitle('Input error');
        $objValidation->setError('Value is not in list.');
        $objValidation->setFormula1('=INDIRECT($A$1)');
        $objPHPExcel->getActiveSheet()->getCell('B1')->setDataValidation($objValidation);


        $path = $_SERVER['DOCUMENT_ROOT'].'/uploads/';
        $objWriter = new \PHPExcel_Writer_Excel2007($objPHPExcel);
        $name = strtotime(Carbon::now());
        $name = $name.'.xlsx';

        $objWriter->save($path.$name);
        chmod($path.$name,0777);

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
