<?php

/**
 * Class App_Util_OfficeDocuments_Document абстрактный класс документа
 */
abstract class App_Util_OfficeDocuments_Document implements App_Util_OfficeDocuments_IDocument
{
    /**
     * @var int Флаг создания документа
     */
    protected $callCreate = 0;
    /**
     * @var ZipArchive основной класс для работы
     */
    protected $zip;

    /**
     * @var string основной путь к папке с рабочим документом
     */
    private $path;

    /**
     * @var string путь к шаблоном
     */
    protected $pathTemplate;

    /**
     * @var string путь к стандартным шаблонам
     */
    protected $pathDefaultTemplate;

    /**
     * @var string путь к шаблоном создания документа
     */
    protected $pathCreateTemplate;

    /**
     * @var string путь к пользовательским шаблонам
     */
    protected $pathCustomTemplate;

    /**
     * @var string путь к временной папке
     */
    protected $pathTemp;

    /**
     * @var string имя файла
     */
    protected $realFileName;

    /**
     * @var string временное имя файла
     */
    protected $tempFileName;

    /**
     * @var mixed Содержимое документа
     */
    protected $content;

    /**
     * @var int Множитель для перевода размеров изображений из пикселей в EMU
     */
    protected $px_emu = 8625;

    /**
     * @var array Делаем приватно, чтобы не было возможности вшить дрянь в документ
     */
    protected $rels = array();

    protected $tempFile = null;

    /**
     * App_Util_OfficeDocuments_Document constructor.
     * @param $filename string имя файла
     * @param string $path_user_template путь к пользовательским шаблонам
     * @throws App_Util_OfficeDocuments_Exception
     */
    public function __construct($filename, $path_user_template = 'user/template/')
    {
        $this->zip = new ZipArchive();
        $this->pathTemplate = $this->getRootPath().'template'.DIRECTORY_SEPARATOR;
        $this->pathTemp = $this->getRootPath().'temp'.DIRECTORY_SEPARATOR;

        $this->tempFileName = $this->generateName() . '.' . $this->expansion;

        $this->realFileName = $filename;

        $this->pathCreateTemplate = $this->pathTemplate  . 'create'   . DIRECTORY_SEPARATOR;
        $this->pathCustomTemplate = $this->pathTemplate  . 'custom'   . DIRECTORY_SEPARATOR;
        $this->pathDefaultTemplate = $this->pathTemplate . 'default'  . DIRECTORY_SEPARATOR;

        // Если не получилось открыть файл, то жизнь бессмысленна.
        if ( $this->zip->open( $this->pathTemp.$this->tempFileName, ZIPARCHIVE::CREATE ) !== TRUE ) {
            throw new App_Util_OfficeDocuments_Exception( "Файл <$filename> не существует\n" );
        }

        // Описываем связи для документа MS Office
        $this->rels = array_merge( $this->rels, array(
            'rId3' => array(
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
                'docProps/app.xml' ),
            'rId2' => array(
                'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
                'docProps/core.xml' ),
        ) );

        // Добавляем типы данных
        $this->zip->addFile($this->pathCreateTemplate . "[Content_Types].xml" , "[Content_Types].xml" );

    }

    // Генерация зависимостей
    protected function add_rels( $filename, $rels, $path = '' ){
        // Шапка XML
        $xmlstring = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        // Добавляем документы по описанным связям
        foreach( $rels as $rId => $params ){
            // Если указан путь к файлу, берем. Если нет, то берем из репозитория
            $pathfile = empty( $params[2] ) ? $this->pathCreateTemplate . $path . $params[1] : $params[2];
            // Добавляем документ в архив
            if( $this->zip->addFile( $pathfile ,  $path . $params[1] ) === false )
                throw new App_Util_OfficeDocuments_Exception('Не удалось добавить в архив ' . $path . $params[1] );
            // Прописываем в связях
            $xmlstring .= '<Relationship Id="' . $rId . '" Type="' . $params[0] . '" Target="' . $params[1] . '"/>';
        }
        $xmlstring .= '</Relationships>';
        // Добавляем в архив
        $this->zip->addFromString( $path . $filename, $xmlstring );
    }

    /**
     *
     */
    private function openTempFile() {

        $flag = 'a';

        if(is_null($this->tempFile) || !file_exists($this->tempFile)) {

            $this->tempFile = $this->pathTemp.$this->generateName().'.bin';
            $flag = 'w';
        }

        $handle = fopen($this->tempFile, $flag);

        return $handle;
    }

    public function addInDocument($content) {

        $fp = $this->openTempFile();

        $write = fwrite($fp, $content); // Запись в файл

        fclose($fp);

        if(!$write)
            throw new App_Util_OfficeDocuments_Exception('Не удалось записать данные в файл');

        return $this->tempFile;
    }

    /**
     * Зменяет все условные переменные на пользовательские данные
     * @param $replace array условнык переменные и пользовательские данные
     * @param $content string контент
     * @return mixed
     */
    public function pparse(array $replace, $content ){
        return str_replace( array_keys( $replace ), array_values( $replace ), $content );
    }

    protected function alreadyCreatedDocument() {
        if($this->callCreate)
            throw new App_Util_OfficeDocuments_Exception('Документ уже был создан, дальнейшее его редактирования невозможно');
    }

    public function setContent($content) {
        $this->content = $content;
    }

    /**
     * Вставка данных в документ
     * @param $var array условнык переменные и пользовательские данные
     * @param $name_template string имя шаблонов
     * @param bool $return
     * @return mixed
     * @throws App_Util_OfficeDocuments_Exception
     */
    protected function assign(array $var, $name_template, $return = false ) {

        // По умолчанию ищем в пользовательских шаблонах
        $path = $this->pathCustomTemplate;

        // Если шаблона нет, то жизнь не имеет смысла
        if(!file_exists( $this->pathDefaultTemplate.$name_template.'.xml' ) &&
            !file_exists( $this->pathCustomTemplate.$name_template.'.xml' ) ) {

            throw new App_Util_OfficeDocuments_Exception( "Пользовательский шаблон <$name_template> не найден!" );
        }

        // Если шаблона нет в пользовательских, но есть шаблон по умолчанию используем его
        if(!file_exists( $this->pathCustomTemplate.$name_template.'.xml' ) && file_exists( $this->pathDefaultTemplate.$name_template.'.xml' ))
            $path = $this->pathDefaultTemplate;

        // Берем шаблон абзаца
        $block = file_get_contents( $path.$name_template.'.xml' );
        $xml = $this->pparse( $var, $block );

        // Если нам указали, что нужно возвратить XML, возвращаем
        if( $return )
            return $xml;
        else
            $this->content .= $xml;

        return true;
    }

    /**
     * Возвращает путь к пользовательским шаблонам
     * @return string
     */
    public function getCustomTemplatePath() {
        return $this->pathCustomTemplate;
    }

    public function getDefaultTemplatePath() {
        return $this->pathDefaultTemplate;
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

    public function create($my_file_document = null) {
        $this->callCreate = true;
    }

    public function __toString()
    {
        return $this->content ? $this->content : 'Документ пустой';
    }

    public function getInfo() {

        return array(
            'tempFile' => $this->pathTemp.$this->tempFileName,
            'name' => $this->realFileName,
            'expansion' => $this->expansion
        );

    }

    /*public function __destruct()
    {
        if(!$this->callCreate)
            $this->create();

        if(file_exists($this->pathTemp.$this->tempFileName))
            unlink($this->pathTemp.$this->tempFileName);
    }*/
}