<?php

namespace App\Livewire;

use Livewire\Component;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\TemplateProcessor;
use PhpOffice\PhpWord\Settings;

use Exception;

class Phpoffice extends Component
{
    public function generateDocx()
    {
        $phpWord = new PhpWord();
        $section = $phpWord->addSection();
        $text = 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod';
        $section->addText($text);
        $objWriter = IOFactory::createWriter($phpWord, 'Word2007');
        try {
            $objWriter->save(storage_path('/app/public/helloWorld.docx'));
        } catch (Exception $e) {

        }

        return response()->download(storage_path('/app/public/helloWorld.docx'));

    }

    //

    public function generateTemplate()
    {
        $templateProcessor = new TemplateProcessor(storage_path('/app/public/template.docx'));
        $templateProcessor->setValue('firstname', 'Sohail');
        $templateProcessor->setValue('lastname', 'Saleem');
        $templateProcessor->saveAs(storage_path('/app/public/result.docx'));

        return response()->download(storage_path('/app/public/result.docx'));
    }

    //

    public function generatePDF()
    {
        $domPdfPath = base_path('vendor/dompdf/dompdf');
        Settings::setPdfRendererPath($domPdfPath);
        Settings::setPdfRendererName('DomPDF');
        $Content = IOFactory::load(storage_path('/app/public/helloWorld.docx'));
        $PDFWriter = IOFactory::createWriter($Content,'PDF');
        $PDFWriter->save(storage_path('/app/public/resultPDF.pdf'));

        return response()->download(storage_path('/app/public/resultPDF.pdf'));
    }

//////////////////////
    public function render()
    {
        return view('livewire.phpoffice');
    }
}
