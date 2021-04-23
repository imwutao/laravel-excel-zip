<?php


namespace Cblink\ExcelZip;


use Carbon\Carbon;
use Chumper\Zipper\Zipper;
use Illuminate\Support\Collection;
use Maatwebsite\Excel\Excel;

class ExcelZip
{

    /**
     * excel instance
     *
     * @var Excel
     */
    private $excel;

    /**
     * folder name
     *
     * @var string
     */
    private $folder;
    private $fileName;

    /**
     * excel export
     *
     * @var
     */
    private $export;

    /**
     * excel counter
     *
     * @var int
     */
    private $counter = 1;

    public function __construct(Excel $excel)
    {
        $this->excel = $excel;
    }

    /**
     * set a excel export
     *
     * @param $export
     * @return $this
     */
    public function setExport($export, $folder)
    {
        $this->export = $export;
        
        $this->folder = $folder.'_'.date('YmdHis');
        $this->fileName = $folder;

        return $this;
    }

    /**
     * store a excel file
     *
     * @param Collection $collection
     * @param string $fileName
     * @param null $export
     * @return $this
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function excel(Collection $collection, $export = null)
    {
        $this->export = $export ?: $this->export;
        $fileName = $this->fileName.'_'.$this->counter;

        $this->excel->store($this->export->setCollection($collection), config('excel_zip.excel_path')."{$this->folder}/$fileName.xlsx");

        $this->counter++;

        return $this;
    }

    /**
     * download zip
     *
     * @return \Symfony\Component\HttpFoundation\BinaryFileResponse
     * @throws \Exception
     */
    public function zip()
    {
        $this->generateZip();

        return $this->response(storage_path('app/'.config('excel_zip.zip_path'))."{$this->folder}.zip");
    }

    /**
     * @param Collection $collection
     * @param CustomCollection $export
     * @param string $fileName
     * @return \Symfony\Component\HttpFoundation\BinaryFileResponse
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @throws \Exception
     */
    /**
     * generate the zip file
     *
     * @throws \Exception
     */
    private function generateZip()
    {
        $zipper = new Zipper();
        
        $zipper->make(storage_path('app/'.config('excel_zip.zip_path'))."{$this->folder}.zip")->add(glob(storage_path('app/'.config('excel_zip.excel_path')).$this->folder.'/*'));

        dispatch(new RemoveZip($this->folder))->delay(Carbon::now()->addMinute());
    }

    /**
     * return response
     *
     * @param string $path
     * @return \Symfony\Component\HttpFoundation\BinaryFileResponse
     */
    private function response(string $path)
    {
        $this->reset();

        return response()->download($path)->deleteFileAfterSend(true);
    }

    /**
     * reset folder
     */
    private function reset()
    {
        $this->folder = null;
        $this->counter = 1;
    }

}