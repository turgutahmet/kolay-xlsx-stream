<?php

namespace Kolay\XlsxStream\Tests;

use Orchestra\Testbench\TestCase as Orchestra;
use Kolay\XlsxStream\XlsxStreamServiceProvider;
use Dotenv\Dotenv;

class TestCase extends Orchestra
{
    /**
     * Setup the test environment.
     */
    protected function setUp(): void
    {
        // Load .env file before parent setup
        // This allows tests to access AWS credentials and other env vars
        $this->loadEnvironmentVariables();

        parent::setUp();
    }

    /**
     * Load environment variables from .env file
     */
    protected function loadEnvironmentVariables(): void
    {
        $envFile = dirname(__DIR__) . '/.env';

        if (file_exists($envFile)) {
            // This is safe in test environment and allows .env to be loaded properly
            $dotenv = Dotenv::createUnsafeImmutable(dirname(__DIR__));
            $dotenv->load();
        }
    }

    /**
     * Get package providers.
     */
    protected function getPackageProviders($app)
    {
        return [
            XlsxStreamServiceProvider::class,
        ];
    }
}