<?php

namespace Kolay\XlsxStream\Tests;

use Orchestra\Testbench\TestCase as Orchestra;
use Kolay\XlsxStream\XlsxStreamServiceProvider;

class TestCase extends Orchestra
{
    protected function getPackageProviders($app)
    {
        return [
            XlsxStreamServiceProvider::class,
        ];
    }
}