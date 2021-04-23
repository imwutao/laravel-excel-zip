<?php

return [
    /**
     * if always_zip is true, when data are less than 5000, it will output as a zip;
     * if false, it will be just one xlsx
     */
    'always_zip' => false,

    'zip_path' => 'exports/',
    
    'excel_path' => 'exports/',

    /**
     * each excel data number
     */
    'chunk' => 100000,

    /**
     * Queue Driver Configuration.
     */
    'queue' => [
        'connection' => null,
    ],
];
