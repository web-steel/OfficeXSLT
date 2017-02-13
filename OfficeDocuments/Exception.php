<?php

/**
 * Created by IntelliJ IDEA.
 * User: HP
 * Date: 28.07.2016
 * Time: 15:31
 */
class App_Util_OfficeDocuments_Exception extends Exception
{
    public static function errorHandlerCallback($code, $string, $file, $line)
    {
        $e = new self($string, $code);
        $e->line = $line;
        $e->file = $file;
        throw $e;
    }
}