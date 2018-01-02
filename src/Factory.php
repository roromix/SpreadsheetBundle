<?php

namespace Roromix\Bundle\SpreadsheetBundle;

use Symfony\Component\HttpFoundation\StreamedResponse;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;
use PhpOffice\PhpSpreadsheet\Helper\Html;

/**
 * Factory for Spreadsheet, StreamedResponse, and IWriter.
 *
 * @package Roromix\Bundle\SpreadsheetBundle
 */
class Factory
{
    private $phpSpreadsheetIO;

    public function __construct($phpSpreadsheetIO = IOFactory::class)
    {
        $this->phpSpreadsheetIO = $phpSpreadsheetIO;
    }

    /**
     * Creates an empty Spreadsheet Object if the filename is empty, otherwise loads the file into the object.
     *
     * @param string $filename
     *
     * @return Spreadsheet
     */
    public function createSpreadsheet($filename = null)
    {
        return (null === $filename) ? new Spreadsheet() : call_user_func(array($this->phpSpreadsheetIO, 'load'), $filename);
    }

    /**
     * Create a worksheet drawing
     * @return Drawing
     */
    public function createSpreadsheetWorksheetDrawing()
    {
        return new Drawing();
    }

    /**
     * Create a writer given the Spreadsheet Object and the type, the type could be one of IOFactory::$writers
     *
     * @param Spreadsheet $spreadsheet
     * @param string $type
     *
     * @return IWriter
     */
    public function createWriter(Spreadsheet $spreadsheet, $type = 'Xls')
    {
        return call_user_func(array($this->phpSpreadsheetIO, 'createWriter'), $spreadsheet, $type);
    }

    /**
     * Stream the file as Response.
     *
     * @param IWriter $writer
     * @param int $status
     * @param array $headers
     *
     * @return StreamedResponse
     */
    public function createStreamedResponse(IWriter $writer, $status = 200, $headers = array())
    {
        return new StreamedResponse(
            function () use ($writer) {
                $writer->save('php://output');
            },
            $status,
            $headers
        );
    }

    /**
     * Create a PHPExcel Helper HTML Object
     *
     * @return Html
     */
    public function createHelperHTML()
    {
        return new Html();
    }
}
