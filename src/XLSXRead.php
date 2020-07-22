<?PHP
/**
 * Created by PhpStorm.
 * User: Fuwa
 * Date: 2019/05/24
 * Time: 14:50
 */

namespace Fuwa;

class XLSXRead{
    private $sPath;
    private $sDir = '';
    private $sError = '';
    private $bError = true;
    private $aSheetName = [];
    private $sAttributesName = '@attributes'; // 节点名称
    private $aStrings;
    private $aSheet = [];

    private $bReadSheet = false; // 是否只读第一个sheet
    private $aReadSheet = []; // 读取哪几个sheet
    private $aReadKey = []; // 读取的key值
    private $bReadKey = false;
    private $aReadFormat = []; // 设置标准
    private $oCallback;

    private $iTitleRows = 1; // 多重标题头,一般为1
    private $aSheetTitle = [];
    private $aTitleSheet = [];
    private $originTitle = [];//原始表头

    private $aConfig = []; // 调用者的配置文件

    public function __construct() {
        $this->bError = true;
    }

    /**
     * @return array
     */
    public function getOriginTitle()
    {
        return $this->originTitle;
    }


    public function read($_aConfig) {
        $this->setConfig($_aConfig);
        $this->setPath();
        if (!$this->bError){
            $this->deleteDir($this->sDir);
            return false;
        }
        $this->setSheetName();
        $this->setSharedStrings();

        $this->aReadSheet = $this->aReadSheet ?: ($this->bReadSheet ? [current(array_flip($this->aSheetName))] : array_flip($this->aSheetName));

        $this->setTitle();
        if (!$this->bError){
            $this->deleteDir($this->sDir);
            return false;
        }

        $this->setContent();

        $this->deleteDir($this->sDir);
        return $this->bError;
    }

    public function setPath($_sPath = false) {
        $this->sPath = $_sPath ?: $this->sPath;
        $this->sDir  = $this->sPath . '.temp.dir';
        if (!$this->error(strtolower(pathinfo($this->sPath, PATHINFO_EXTENSION)) != 'xlsx', '只支持Xlsx 2007')) {
            return false;
        }


        $zip = new \ZipArchive();
        if ($zip->open($this->sPath) === true) {
            for ($i = 0; $i < $zip->numFiles; $i++) {
                $filename = $zip->getNameIndex($i);
                $fileinfo = pathinfo($filename);
                if (!is_dir($this->sDir . '/' . $fileinfo['dirname'])) {
                    @mkdir($this->sDir . '/' . $fileinfo['dirname'], 0777, true);
                }
                copy("zip://" . $this->sPath . "#" . $filename, $this->sDir . '/' . $filename);
            }
            $zip->close();
        }

        if (!$this->error(!is_file($this->sDir . '/xl/workbook.xml'), '只支持Xlsx-2007')) {
            return false;
        }

        return true;
    }

    private function setConfig($_aConfig) {
        $this->aConfig     = $_aConfig;
        $this->sPath       = $this->aConfig['path'];
        $this->bReadSheet  = isset($this->aConfig['bSheet']) ? $this->aConfig['bSheet'] : $this->bReadSheet;
        $this->aReadSheet  = isset($this->aConfig['aSheet']) ? (array)$this->aConfig['aSheet'] : $this->aReadSheet;
        $this->aReadKey    = isset($this->aConfig['aKey']) ? (array)$this->aConfig['aKey'] : $this->aReadKey;
        $this->aReadFormat = isset($this->aConfig['format']) ? (array)$this->aConfig['format'] : $this->aReadFormat;
        $this->oCallback   = isset($this->aConfig['callback']) ? $this->aConfig['callback'] : false;
        $this->iTitleRows  = isset($this->aConfig['titleRows']) ? $this->aConfig['titleRows'] : $this->iTitleRows;
    }
    private function setSheetName() {
        $aData = $this->xmlToArray($this->sDir . '/xl/workbook.xml');
        foreach ($aData['sheets']['sheet'] as $_k => $_v) {
            $this->aSheetName[isset($_v['name']) ? $_v['name'] : $_v[$this->sAttributesName]['name']] = 'sheet' . ($_k + 1);
        }
        unset($aData);
    }
    private function setSharedStrings(){
        $this->aStrings = $this->xmlToArray($this->sDir . '/xl/sharedStrings.xml');
    }
    private function setTitle(){
        foreach ($this->aReadSheet as $_v){
            $this->aSheet[$_v]['sXmlPath']     = $this->sDir . '/xl/worksheets/' . $this->aSheetName[$_v] . '.xml';
            if (!is_file($this->aSheet[$_v]['sXmlPath'])) {
                $this->error(true, "【{$_v}】不存在！");
                return false;
            }
            $this->aSheet[$_v]['sNoteXmlPath'] = $this->aSheet[$_v]['sXmlPath'] . '.xml';
            $streamer = \Prewk\XmlStringStreamer::createStringWalkerParser($this->aSheet[$_v]['sXmlPath']);
            while ($node = $streamer->getNode()) {
                if (strpos(trim($node), '<sheetData>') === 0) {
                    $this->aSheetTitle[$_v] = [];
                    file_put_contents($this->aSheet[$_v]['sNoteXmlPath'], '<?xml version="1.0" encoding="UTF-8"?>' . PHP_EOL . $node);
                    break;
                }
            }

            if (!isset($this->aSheetTitle[$_v]))
                continue;

            $this->aSheet[$_v]['oNode'] = \Prewk\XmlStringStreamer::createStringWalkerParser($this->aSheet[$_v]['sNoteXmlPath']);

            while ($node_ = $this->aSheet[$_v]['oNode']->getNode()) {
                $node = $this->xmlToArray($node_);
                $row = $node[$this->sAttributesName]['r'];

                if ($row > $this->iTitleRows) {
                    $this->aSheet[$_v]['sNode'][] = $node_;
                    break;
                }

                isset($node['c'][$this->sAttributesName]) && $node['c'][0] = $node['c'];

                foreach ($node['c'] as $_cv) {
                    $srow = $_cv[$this->sAttributesName]['r'];

                    if (!isset($_cv['v']) && !isset($_cv['is']['t'])) continue;
                    $_cv['v'] = isset($_cv['v']) ? $_cv['v'] : $_cv['is']['t'];

                    $sContent = isset($_cv[$this->sAttributesName]['t']) && !isset($_cv['is']) ? $this->getStrings($_cv['v']) : $_cv['v'];
                    $sContent = str_replace(PHP_EOL, '', $sContent);

                    $this->aSheetTitle[$_v][rtrim($srow, $row)] = $sContent;

                    $this->originTitle[$_v][rtrim($srow, $row)] = $sContent;//存放原始表头
                }
            }

            if (!empty($this->aReadKey)){
                $aSheetTitle = [];
                foreach ($this->aReadKey as $_readk => $_readv){
                    if (is_string($_readv)){
                        $zh = array_search($_readv,$this->aSheetTitle[$_v]);
                        false !== $zh && $aSheetTitle[$_readk] = $zh;
                    }else{
                        foreach ($_readv as $_readv_v){
                            $zh = array_search($_readv_v,$this->aSheetTitle[$_v]);
                            false !== $zh && $aSheetTitle[$_readk] = $zh;
                        }
                    }
                }
                $this->aSheetTitle[$_v] = $aSheetTitle;
                $this->bReadKey = true;
            }
        }

        $aEqualKey = current($this->aSheetTitle);
        foreach ($this->aSheetTitle as $_k => $_v){
            if (!$this->judgeEqualKey($aEqualKey, $_v)){
                $this->error(true, '多个sheet标题不相符');
                return false;
            }
            $this->aTitleSheet[$_k] = array_flip($_v);
        }

    }
    private function setContent(){
        $aCon = ['aTitleSheet' => $this->aTitleSheet];
        foreach ($this->aTitleSheet as $_k => $_v){
            while ($node = $this->getNode($_k)){
                $node = $this->xmlToArray($node);
                $row = $node[$this->sAttributesName]['r'];
                $aContent = [];

                $node['c'] = isset($node['c'][$this->sAttributesName]) ? [$node['c']] : $node['c'];
                foreach ($node['c'] as $_nodev) {
                    $srow    = $_nodev[$this->sAttributesName]['r'];
                    $rowname = rtrim($srow, $row);

                    if (((!isset($_nodev['v']) && !isset($_nodev['is']['t'])) || ($this->bReadKey && !array_key_exists($rowname,$_v))) && ($this->bReadKey && !array_key_exists($rowname,$_v)))
                        continue;

                    $_nodev['v'] = isset($_nodev['v']) ? $_nodev['v'] : $_nodev['is']['t'] ?: '';
                    $aContent[$this->bReadKey ? $_v[$rowname] : $rowname] = isset($_nodev[$this->sAttributesName]['t']) && !isset($_nodev['is']) ? ($this->getStrings($_nodev['v']) === NULL ? $_nodev['v'] : $this->getStrings($_nodev['v'])) : $this->getFormat($_nodev, $this->bReadKey ? $_v[$rowname] : false);
                }

                if ($oCallBack = $this->oCallback){
                    $aCon['sheet'] = $_k;
                    $flag = $oCallBack($aContent, $row, $aCon); // call_user_func

                    if ($flag === false){
                        return false;
                    }elseif ($flag == 'continue'){
                        continue;
                    }elseif ($flag == 'break'){
                        break;
                    }elseif ($flag == 'continue 2'){
                        continue 2;
                    }elseif ($flag == 'break 2'){
                        break 2;
                    }elseif (is_array($flag) && $flag[0] === false){
                        $this->sError = $flag[1];
                        $this->bError = false;
                        return false;
                    }
                }

            }
        }
    }

    private function getNode($_sSheetName){
        if (!empty($this->aSheet[$_sSheetName]['sNode'])){
            return array_shift($this->aSheet[$_sSheetName]['sNode']);
        }else{
            return $this->aSheet[$_sSheetName]['oNode']->getNode();
        }
    }
    public function judgeEqualKey($_aKey1, $_aKey2) {
        if (array_diff_key($_aKey1, $_aKey2) || array_diff_key($_aKey2, $_aKey1)) {
            return false;
        } else {
            return true;
        }
    }
    public function getStrings($_iIndex) {
        if (isset($this->aStrings['si'][$_iIndex])) {
            if (is_string($this->aStrings['si'][$_iIndex]['t'])) {
                return (string)$this->aStrings['si'][$_iIndex]['t'];
            } else {
                return (string)implode('', array_column($this->aStrings['si'][$_iIndex]['r'], 't'));
            }
        }
    }
    public function getFormat($_aArr,$name = false){
        $arr = $_aArr;
        $content = $arr['v'];
        if (isset($name,$this->aReadFormat)) {
            switch ($this->aReadFormat[$name]) {
                case 'datetime':
                    $content = empty($arr['v']) ? '' : date('Y-m-d H:i:s', $arr['v'] * 86400 - 2209190400);
                    break;
                case 'time':
                    $content = empty($arr['v']) ? 0 : $arr['v'] * 86400 - 2209190400;
                    break;
            }
        }
        return (string)$content;
    }
    private function error($_bIs, $_sError) {
        if (!$_bIs)
            return true;

        $this->bError = false;
        $this->sError = $_sError;
        return false;
    }
    private function xmlToArray($xml) {
        try {
            $xml = str_replace('gbk', 'UTF-8', $xml);
            $xml = str_replace('gb2312', 'UTF-8', $xml);
            $xml = str_replace('GBK', 'UTF-8', $xml);
            $xml = str_replace('GB2312', 'UTF-8', $xml);
            return json_decode(json_encode(is_file($xml) ? simplexml_load_file($xml, 'SimpleXMLElement', LIBXML_NOCDATA) : simplexml_load_string($xml)), true);
        } catch (\Exception $exception) {
            $this->bError = false;
            $this->sError = $exception->getMessage();
            return $this->bError;
        }
    }
    public function __get($name) {
        // TODO: Implement __get() method.
        return $this->$name;
    }
    public function getMessage(){
        return $this->sError;
    }
    function deleteDir($dir) {
        if (!$handle = @opendir($dir)) {
            return false;
        }
        while (false !== ($file = readdir($handle))) {
            if ($file !== "." && $file !== "..") {
                $file = $dir . '/' . $file;
                if (is_dir($file)) {
                    $this->deleteDir($file);
                } else {
                    @unlink($file);
                }
                @rmdir($dir);
            }
        }
    }
}