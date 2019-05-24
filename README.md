# XLSXRead
###Simply read large files 
```PHP
$oXLSXRead = new \Fuwa\XLSXRead();
$oCallback = function ($_v, $row, $con) {
    if ($row < 6)
        return 'continue';
    if (empty($_v['id']))
        return [false, 'Number cannot be empty!'];
};
$aConfig = ['path'      => $this->getTempCacheDir() . 'csv/sf.xlsx',
            'aKey'      => ['id' => 'Serial number', 'comtime' => 'Completion time',],
            'format'    => ['comtime' => 'time'],
            'aSheet'    => ['progress'],
            'callback'  => $oCallback,
            'titleRows' => 2,];
$flag = $oXLSXRead->read($aConfig);
if (false === $flag){
    $oXLSXRead->getMessage();
}