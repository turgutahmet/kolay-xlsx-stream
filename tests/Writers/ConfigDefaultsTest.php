<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Sinks\S3MultipartSink;
use Kolay\XlsxStream\Tests\TestCase;
use Kolay\XlsxStream\Writers\BaseXlsxWriter;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

/**
 * v3.2.2 config semantics: the published config/xlsx-stream.php is
 * applied at writer construction — version-gated so pre-3.2.2
 * published copies (whose keys were never read and whose stale values
 * contradict the code defaults) stay inert, and always overridable by
 * code-level setters.
 *
 * The probe subclass exposes the protected fold + the resulting state;
 * no global config() shim is defined (the suite runs under Testbench,
 * whose real config() the constructor path already exercises).
 */
class ConfigDefaultsTest extends TestCase
{
    protected function setUp(): void
    {
        parent::setUp();

        // Neutralize the package's own published config (merged by the
        // service provider) so every probe starts from the pure code
        // defaults and each scenario feeds applyConfigDefaults()
        // exactly the array it means to test.
        config(['xlsx-stream' => []]);
    }

    private function probe(): ConfigProbeWriter
    {
        return new ConfigProbeWriter;
    }

    public function test_missing_version_key_keeps_code_defaults(): void
    {
        $w = $this->probe();
        $w->applyConfig(['writer' => ['compression_level' => 9, 'buffer_flush_interval' => 5]]);

        $this->assertSame(5, $w->deflateLevel());
        $this->assertSame(1000, $w->bufferFlushInterval());
    }

    public function test_version_1_config_is_ignored(): void
    {
        $w = $this->probe();
        $w->applyConfig(['version' => 1, 'writer' => ['compression_level' => 9]]);

        $this->assertSame(5, $w->deflateLevel());
    }

    public function test_version_2_compression_level_applies(): void
    {
        $w = $this->probe();
        $w->applyConfig(['version' => 2, 'writer' => ['compression_level' => 7]]);

        $this->assertSame(7, $w->deflateLevel());
    }

    public function test_setter_overrides_config(): void
    {
        $w = $this->probe();
        $w->applyConfig(['version' => 2, 'writer' => ['compression_level' => 7]]);
        $w->setCompressionLevel(3);

        $this->assertSame(3, $w->deflateLevel());
    }

    public function test_invalid_values_are_silently_ignored(): void
    {
        $w = $this->probe();
        $w->applyConfig(['version' => 2, 'writer' => [
            'compression_level' => 0,          // out of range
            'buffer_flush_interval' => 'abc',  // not numeric
        ]]);

        $this->assertSame(5, $w->deflateLevel());
        $this->assertSame(1000, $w->bufferFlushInterval());

        $w->applyConfig(['version' => 2, 'writer' => ['compression_level' => 'abc']]);
        $this->assertSame(5, $w->deflateLevel());
    }

    public function test_buffer_flush_interval_applies(): void
    {
        $w = $this->probe();
        $w->applyConfig(['version' => 2, 'writer' => ['buffer_flush_interval' => 5000]]);

        $this->assertSame(5000, $w->bufferFlushInterval());
    }

    public function test_null_config_is_a_no_op(): void
    {
        $w = $this->probe();
        $w->applyConfig(null);

        $this->assertSame(5, $w->deflateLevel());
        $this->assertSame(1000, $w->bufferFlushInterval());
    }

    public function test_constructor_applies_live_laravel_config(): void
    {
        config(['xlsx-stream' => ['version' => 2, 'writer' => ['compression_level' => 8]]]);

        $this->assertSame(8, $this->probe()->deflateLevel());
    }

    public function test_constructor_ignores_pre_322_shaped_config(): void
    {
        // The exact shape a pre-3.2.2 published copy has: no version
        // key, stale level 1. Must stay as inert as it always was.
        config(['xlsx-stream' => ['writer' => ['compression_level' => 1]]]);

        $this->assertSame(5, $this->probe()->deflateLevel());
    }

    // ------------------------------------------------------------------
    // forDisk() part-size resolution (pure helper — no S3 client needed)
    // ------------------------------------------------------------------

    private function resolvePartSize(?int $explicit, ?array $config): int
    {
        $m = new \ReflectionMethod(SinkableXlsxWriter::class, 'resolvePartSize');

        return $m->invoke(null, $explicit, $config);
    }

    public function test_part_size_explicit_argument_wins(): void
    {
        $this->assertSame(
            16 * 1024 * 1024,
            $this->resolvePartSize(16 * 1024 * 1024, ['version' => 2, 's3' => ['part_size' => 5242880]])
        );
    }

    public function test_part_size_from_version_2_config(): void
    {
        $this->assertSame(
            33554432,
            $this->resolvePartSize(null, ['version' => 2, 's3' => ['part_size' => 33554432]])
        );
    }

    public function test_part_size_ignores_unversioned_config(): void
    {
        $this->assertSame(
            S3MultipartSink::DEFAULT_PART_SIZE,
            $this->resolvePartSize(null, ['s3' => ['part_size' => 33554432]])
        );
    }

    public function test_part_size_defaults_without_config(): void
    {
        $this->assertSame(S3MultipartSink::DEFAULT_PART_SIZE, $this->resolvePartSize(null, null));
        $this->assertSame(
            S3MultipartSink::DEFAULT_PART_SIZE,
            $this->resolvePartSize(null, ['version' => 2, 's3' => ['part_size' => 'abc']])
        );
    }

    // ------------------------------------------------------------------
    // forDisk() concurrency resolution (pure helper — no S3 client needed)
    // ------------------------------------------------------------------

    private function resolveConcurrency(?array $config): int
    {
        $m = new \ReflectionMethod(SinkableXlsxWriter::class, 'resolveConcurrency');

        return $m->invoke(null, $config);
    }

    public function test_concurrency_defaults_without_version_key(): void
    {
        $this->assertSame(
            S3MultipartSink::DEFAULT_CONCURRENCY,
            $this->resolveConcurrency(['s3' => ['concurrency' => 8]])
        );
        $this->assertSame(S3MultipartSink::DEFAULT_CONCURRENCY, $this->resolveConcurrency(null));
    }

    public function test_concurrency_from_version_2_config(): void
    {
        $this->assertSame(8, $this->resolveConcurrency(['version' => 2, 's3' => ['concurrency' => 8]]));
    }

    public function test_concurrency_ignores_invalid_values(): void
    {
        foreach ([0, -1, 'x'] as $invalid) {
            $this->assertSame(
                S3MultipartSink::DEFAULT_CONCURRENCY,
                $this->resolveConcurrency(['version' => 2, 's3' => ['concurrency' => $invalid]]),
                'invalid value: '.var_export($invalid, true)
            );
        }
    }

    public function test_concurrency_ignores_version_1_config(): void
    {
        $this->assertSame(
            S3MultipartSink::DEFAULT_CONCURRENCY,
            $this->resolveConcurrency(['version' => 1, 's3' => ['concurrency' => 8]])
        );
    }
}

/**
 * Minimal concrete writer: no-op destination, protected state exposed.
 */
class ConfigProbeWriter extends BaseXlsxWriter
{
    protected function writeToDest(string $data): void
    {
        // discard — construction/config behaviour is all these tests need
    }

    public function applyConfig(?array $config): void
    {
        $this->applyConfigDefaults($config);
    }

    public function deflateLevel(): int
    {
        return $this->deflateLevel;
    }

    public function bufferFlushInterval(): int
    {
        return $this->bufferFlushInterval;
    }
}
