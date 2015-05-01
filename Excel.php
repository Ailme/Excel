<?php
/**
 *
 * User: Alexandr Tumaykin
 * Date: 18.01.2015
 * Time: 12:22
 *
 */

namespace ailme\excel;


class Excel
{

    //и позиционирование
    private $_alignCenterTop = [
        'alignment' => [
            'horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            'vertical'   => \PHPExcel_Style_Alignment::VERTICAL_TOP
        ]
    ];

    /**
     * @var \PHPExcel_Reader_IReader
     */
    protected $_reader;
    /**
     * @var \PHPExcel_Writer_IWriter
     */
    protected $_writer;
    /**
     * Excel object
     * @var \PHPExcel
     */
    protected $_excel;
    /**
     * @var string
     */
    protected $_ext = '.xlsx';
    /**
     * @var
     */
    protected $_title;
    /**
     * @var array
     */
    protected $_headers = [ ];
    /**
     * @var array
     */
    protected $_columnWidth = [ ];
    /**
     * @var int
     */
    protected $_rowIndex = 1;
    /**
     * @var int
     */
    protected $_rowStart = 1;
    /**
     * @var
     */
    protected $_cells;
    /**
     * @var \PHPExcel_Worksheet
     */
    protected $_sheet;

    /**
     * @var \PHPExcel_Worksheet_Row
     */
    protected $_row;

    /**
     * @var \PHPExcel_Worksheet_CellIterator
     */
    protected $_cellIterator;

    /**
     * @var bool
     */
    protected $_readDataOnly = FALSE;

    /**
     * @var string
     */
    protected $_filename;

    /**
     * используется для экспорта в excel
     *
     * @param array  $headers
     *
     * @param string $a
     *
     * @return array
     */
    public function letterCells( array $headers = [ ], $a = 'A' )
    {
        $result = [ ];

        foreach ( $headers as $i => $header ) {
            //        $result[num2alpha($i)] = $header;
            $result[ $a ] = $header;
            $a++;
        }

        return $result;
    }

    /**
     * @param string   $title
     * @param \Closure $callback
     *
     * @return $this
     */
    public function create( $title, $callback = NULL )
    {
        $this->_title = $title;
        $this->_excel = new \PHPExcel();
        $this->_excel->disconnectWorksheets();
        $this->_writer = \PHPExcel_IOFactory::createWriter( $this->_excel, 'Excel2007' );

        if ( $callback instanceof \Closure ) {
            call_user_func( $callback, $this->_excel );
        }

        return $this;
    }

    /**
     * download file
     */
    public function save()
    {
        header( 'Content-Type: application/vnd.ms-excel' );
        header( 'Content-Disposition: attachment;filename="' . $this->_title . $this->_ext . '"' );
        header( 'Cache-Control: max-age=0' );
        $this->_writer->save( 'php://output' );
        die();
    }

    /**
     * download file
     */
    public function saveMacros( $title = NULL )
    {
        $this->_title = $title ?: $this->_title;

        header( 'Content-Type: application/vnd.ms-excel' );
        header( 'Content-Disposition: attachment;filename="' . $this->_title . '.xlsm"' );
        header( 'Cache-Control: max-age=0' );
        $this->_writer->save( 'php://output' );
        die();
    }

    /**
     * @param string   $title
     * @param \Closure $callback
     *
     * @return \PHPExcel_Worksheet
     */
    public function createSheet( $title = NULL, $callback = NULL )
    {
        $this->_sheet    = $this->_excel->createSheet();
        $this->_rowIndex = $this->_rowStart;

        if ( is_string( $title ) && strlen( $title ) ) {
            $this->_sheet->setTitle( $title );
        }

        if ( $callback instanceof \Closure ) {
            call_user_func( $callback, $this->_sheet );
        }

        return $this->_sheet;
    }

    /**
     * @param array $headers
     * @param array $columnWidth
     *
     * @return $this
     * @throws \PHPExcel_Exception
     */
    public function setHeaders( array $headers, $columnWidth = [ ] )
    {
        $this->_headers = $this->letterCells( $headers );
        $this->_cells   = array_keys( $this->_headers );

        foreach ( $this->_headers as $cell => $value ) {
            //устанавливаем заголовки
            $this->_sheet->setCellValue( $cell . $this->_rowIndex, $value );
            // выравниваем заголовки по центру
            $this->_sheet->getStyle( $cell . $this->_rowIndex )->applyFromArray( $this->_alignCenterTop );
        }

        //устанавливаем ширину
        if ( sizeof( $columnWidth ) ) {
            foreach ( $this->letterCells( $columnWidth ) as $cell => $width ) {
                $this->_sheet->getColumnDimension( $cell )->setWidth( $width );
            }
        }

        $this->_rowIndex++;

        return $this;
    }

    /**
     * @param                     $items
     * @param \Closure            $callback
     *
     * @return $this
     */
    public function addRows( $items, $callback = NULL )
    {
        foreach ( $items as $item ) {
            if ( $callback instanceof \Closure ) {
                call_user_func( $callback, $item, $this->_cells, $this->_rowIndex );
                $this->_rowIndex++;
            }
        }

        return $this;
    }

    /**
     * @param          $cells
     * @param          $row
     * @param \Closure $callback
     *
     * @return $this
     */
    public function addCell( &$cells, $row, $callback = NULL )
    {
        if ( sizeof( $cells ) ) {
            $coordinate = array_shift( $cells ) . $row;
            $cell       = $this->_sheet->getCell( $coordinate );

            if ( $callback instanceof \Closure ) {
                call_user_func( $callback, $cell, $coordinate );
            }
        }

        return $this;
    }

    /**
     * @return $this
     */
    public function setAutoFilter()
    {
        $lastCell = array_pop( $this->_cells );
        $this->_sheet->setAutoFilter( 'A' . $this->_rowStart . ':' . $lastCell . $this->_rowIndex );

        return $this;
    }

    /**
     * @param $value
     *
     * @return $this
     */
    public function setStartRow( $value = 1 )
    {
        $this->_rowStart = $value;

        return $this;
    }


    /*************************************************************
     *
     *************************************************************/

    /**
     * @param          $filename
     * @param \Closure $callback
     *
     * @return $this
     */
    public function load( $filename, $callback = NULL )
    {
        $this->_filename = $filename;
        $this->_reader   = \PHPExcel_IOFactory::createReaderForFile( $this->_filename );

        if ( $callback instanceof \Closure ) {
            $this->_excel  = $this->_reader->setReadDataOnly( $this->_readDataOnly )->load( $this->_filename );
            $this->_writer = \PHPExcel_IOFactory::createWriter( $this->_excel, 'Excel2007' );
            call_user_func( $callback, $this->_excel );
        }

        return $this;
    }

    /**
     * @param int|string $sheet
     * @param \Closure   $callback
     *
     * @return $this
     */
    public function readSheet( $sheet = 0, $callback = NULL )
    {
        $method       = is_numeric( $sheet ) ? 'setActiveSheetIndex' : 'setActiveSheetIndexByName';
        $this->_sheet = $this->_excel->$method( $sheet );

        if ( $callback instanceof \Closure ) {
            call_user_func( $callback, $this->_sheet );
        }

        return $this;
    }

    /**
     * @param \Closure $callback
     *
     * @return $this
     */
    public function eachRows( $callback = NULL )
    {
        foreach ( $this->_sheet->getRowIterator() as $row ) {
            $this->_row          = $row;
            $this->_cellIterator = $this->_row->getCellIterator();

            if ( $callback instanceof \Closure ) {
                call_user_func( $callback, $this->_row );
            }
        }

        return $this;
    }

    /**
     * @param \Closure $callback
     *
     * @return $this
     */
    public function eachCells( $callback = NULL )
    {
        foreach ( $this->_cellIterator as $cell ) {
            if ( $callback instanceof \Closure ) {
                call_user_func( $callback, $cell );
            }
        }

        return $this;
    }

    /**
     * @param \Closure $callback
     *
     * @return $this|mixed
     */
    public function readCell( $callback = NULL )
    {
        $cell = $this->_cellIterator->current();
        $this->_cellIterator->next();

        if ( $callback instanceof \Closure ) {
            return call_user_func( $callback, $cell );
        }

        return $this;
    }

    /**
     * @return $this
     */
    public function stepCell()
    {
        $this->_cellIterator->next();

        return $this;
    }

    /**
     * @return bool
     */
    public function getReadDataOnly()
    {
        return $this->_readDataOnly;
    }

    /**
     * @param bool $value
     *
     * @return $this
     */
    public function setReadDataOnly( $value = FALSE )
    {
        $this->_readDataOnly = $value;
        return $this;
    }

    /**
     * @param $filename
     *
     * @return $this
     */
    public function setFilename( $filename )
    {
        $this->_filename = $filename;
        return $this;
    }

    /**
     * @return string
     */
    public function getFilename()
    {
        return $this->_filename;
    }

    /**
     * @param $title
     *
     * @return $this
     */
    public function setTitle( $title )
    {
        $this->_title = $title;
        return $this;
    }
}