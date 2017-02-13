<?php

/**
 * Class App_Util_OfficeDocuments_WordDocument_Word документ Word
 */
class App_Util_OfficeDocuments_WordDocument_Word extends App_Util_OfficeDocuments_Document
{
    /**
     * @var string расширение файла
     */
    protected $expansion = 'docx';

    /**
     * App_Util_OfficeDocuments_WordDocument_Word constructor.
     * @param string $filename имя файла
     */
    public function __construct( $filename ){

        parent::__construct( $filename );

        $this->word_rels = array(
            "rId3" => array(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
                "settings.xml",
            ),
            "rId2" => array(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                "styles.xml",
            ),
            "rId1" => array(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
                "numbering.xml"
            ),
            "rId6" => array(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                "theme/theme1.xml",
            ),
            "rId5" => array(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
                "fontTable.xml",
            ),
            "rId4" => array(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
                "webSettings.xml",
            )
        );
    }

    /**
     * @param $var
     * @param $name_template
     * @param bool $return_xml
     * @return mixed
     * @throws App_Util_OfficeDocuments_Exception
     */
    public function insertTemplate($var, $name_template, $return_xml = false) {

        $this->alreadyCreatedDocument();

        return $this->assign($var, $name_template, $return_xml);
    }

    /**
     * Создает файл
     * @param string $my_file_document
     * @throws App_Util_OfficeDocuments_Exception
     */
    public function create($my_file_document = null) {

        parent::create();

        $this->rels['rId1'] = array(
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'word/document.xml' );

        // Добавляем связанные документы MS Office
        $this->add_rels( "_rels/.rels", $this->rels );

        // Добавляем связанные документы MS Office Word
        $this->add_rels( "_rels/document.xml.rels", $this->word_rels, 'word/' );

        if(is_null($my_file_document) || !file_exists($my_file_document) ) {
            // Добавляем содержимое
            $this->zip->addFromString("word/document.xml",
                str_replace('{CONTENT}', $this->content, file_get_contents($this->pathCreateTemplate . "word/document.xml")));

        } else {
            $this->zip->addFile($my_file_document, "word/document.xml");
        }

        $this->zip->close();

        if(!is_null($my_file_document) && file_exists($my_file_document) ) {
            unlink($my_file_document);
        }
    }
}