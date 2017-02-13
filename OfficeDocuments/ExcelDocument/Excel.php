<?php

/**
 * Created by IntelliJ IDEA.
 * User: HP
 * Date: 28.07.2016
 * Time: 14:56
 */
class App_Util_OfficeDocuments_ExcelDocument_Excel extends App_Util_OfficeDocuments_ExcelDocument_Abstract implements App_Util_OfficeDocuments_IDocument
{
    protected $expansion = 'xlsx';

    protected $author = 'Abon';
    protected $pathTemp;
    protected $realFileName = '';
    protected $tempFileName;

    public function __construct($filename)
    {
        $this->temp_dir = $this->getRootPath().'temp'.DIRECTORY_SEPARATOR;
        $this->pathTemp = $this->getRootPath().'temp'.DIRECTORY_SEPARATOR;

        $this->tempFileName = $this->generateName() . '.' . $this->expansion;

        $this->realFileName = $filename;

        parent::__construct();
    }

    public function create($my_file_document = null) {
        parent::writeToFile($this->pathTemp.$this->tempFileName);
    }

    public function writeToFile($filename) {
        $this->create($filename);
    }

    public function getInfo() {

        return array(
            'tempFile' => $this->pathTemp.$this->tempFileName,
            'name' => $this->realFileName,
            'expansion' => $this->expansion
        );

    }

    /**
     * Возращает основнуть путь к папке с рабочим документом
     * @return string
     */
    protected function getRootPath() {

        if($this->path)
            return $this->path;

        $array_path = explode('_', get_class($this));
        array_walk($array_path, function(&$data) {
            $data = $data . DIRECTORY_SEPARATOR;
        });

        $path = str_replace($array_path, '', __DIR__.DIRECTORY_SEPARATOR);
        unset($array_path[count($array_path) - 1]);

        $this->path = $path.implode('', $array_path);

        return $this->path;
    }

    /**
     * Генератор уникальных имен
     * @return string уникальное имя
     */
    private function generateName() {

        $id = uniqid();

        $id = base_convert($id, 16, 2);
        $id = str_pad($id, strlen($id) + (8 - (strlen($id) % 8)), '0', STR_PAD_LEFT);

        $chunks = str_split($id, 8);

        $id = array();
        foreach ($chunks as $key => $chunk) {
            if ($key & 1) {  // odd
                array_unshift($id, $chunk);
            } else {         // even
                array_push($id, $chunk);
            }
        }

        return base_convert(implode($id), 2, 36);
    }
}