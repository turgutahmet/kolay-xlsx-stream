<?php

namespace Kolay\XlsxStream;

use Illuminate\Support\ServiceProvider;

class XlsxStreamServiceProvider extends ServiceProvider
{
    /**
     * Register services.
     *
     * @return void
     */
    public function register()
    {
        // Merge config
        $this->mergeConfigFrom(
            __DIR__.'/../config/xlsx-stream.php', 'xlsx-stream'
        );
    }

    /**
     * Bootstrap services.
     *
     * @return void
     */
    public function boot()
    {
        // Publish config
        if ($this->app->runningInConsole()) {
            $this->publishes([
                __DIR__.'/../config/xlsx-stream.php' => config_path('xlsx-stream.php'),
            ], 'xlsx-stream-config');
        }
    }
}