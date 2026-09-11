<?php

namespace Kolay\XlsxStream\Tests\Sources;

use Kolay\XlsxStream\Contracts\Source;
use Kolay\XlsxStream\Contracts\SupportsBoundedStream;
use Kolay\XlsxStream\Contracts\SupportsSuffixRange;
use Kolay\XlsxStream\Exceptions\XlsxReadException;
use Kolay\XlsxStream\Sources\StringSource;
use Kolay\XlsxStream\Tests\TestCase;

/**
 * StringSource — an in-memory Source, the mirror of LocalFileSource for
 * bytes that never touch disk (a template generated inside the same
 * process). Gates: the Source contract's read semantics, the suffix-range
 * bootstrap ZipDirectory uses, stream independence, and the closed guard.
 */
class StringSourceTest extends TestCase
{
    private const BODY = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';

    private function src(string $data = self::BODY): StringSource
    {
        return new StringSource($data);
    }

    public function test_implements_the_source_capability_set(): void
    {
        $s = $this->src();
        $this->assertInstanceOf(Source::class, $s);
        $this->assertInstanceOf(SupportsSuffixRange::class, $s);
        $this->assertInstanceOf(SupportsBoundedStream::class, $s);
    }

    public function test_size_and_range(): void
    {
        $s = $this->src();
        $this->assertSame(strlen(self::BODY), $s->size());
        $this->assertSame('ABCDE', $s->range(0, 5));
        $this->assertSame('789', $s->range(strlen(self::BODY) - 3, 3));
        $this->assertSame('', $s->range(0, 0));
    }

    public function test_range_past_the_end_clamps_like_a_short_read(): void
    {
        // A file source returns what exists; never pads, never throws.
        $s = $this->src();
        $this->assertSame('6789', $s->range(strlen(self::BODY) - 4, 999));
        $this->assertSame('', $s->range(strlen(self::BODY), 10));
        $this->assertSame('', $s->range(strlen(self::BODY) + 50, 10));
    }

    public function test_tail_returns_bytes_and_total_size_in_one_call(): void
    {
        $s = $this->src();
        $this->assertSame(['data' => '6789', 'size' => strlen(self::BODY)], $s->tail(4));
        // Length beyond the source yields the whole content (contract).
        $this->assertSame(['data' => self::BODY, 'size' => strlen(self::BODY)], $s->tail(9999));
        $this->assertSame(['data' => '', 'size' => strlen(self::BODY)], $s->tail(0));
    }

    public function test_stream_from_reads_to_the_end_and_handles_are_independent(): void
    {
        $s = $this->src();
        $a = $s->streamFrom(26);
        $b = $s->streamFrom(0);

        $this->assertSame('0123456789', stream_get_contents($a));
        // Draining $a must not disturb $b.
        $this->assertSame(self::BODY, stream_get_contents($b));
        $this->assertTrue(feof($a));
        fclose($a);
        fclose($b);
    }

    public function test_stream_from_range_is_bounded(): void
    {
        $s = $this->src();
        $h = $s->streamFromRange(26, 4);
        $this->assertSame('0123', stream_get_contents($h));
        fclose($h);
    }

    public function test_empty_source(): void
    {
        $s = $this->src('');
        $this->assertSame(0, $s->size());
        $this->assertSame('', $s->range(0, 10));
        $this->assertSame(['data' => '', 'size' => 0], $s->tail(64));
        $h = $s->streamFrom(0);
        $this->assertSame('', stream_get_contents($h));
        fclose($h);
    }

    public function test_binary_safety(): void
    {
        $bin = "\x00\x01\x02\xFF\xFEbytes\x00";
        $s = $this->src($bin);
        $this->assertSame(strlen($bin), $s->size());
        $this->assertSame($bin, $s->range(0, strlen($bin)));
    }

    public function test_close_is_idempotent_and_guards_reads(): void
    {
        $s = $this->src();
        $s->close();
        $s->close(); // idempotent
        $this->expectException(XlsxReadException::class);
        $s->range(0, 1);
    }
}
