<?php

namespace Kolay\XlsxStream\Tests\Writers;

use Kolay\XlsxStream\Tests\TestCase;

use Kolay\XlsxStream\Exceptions\XlsxStreamException;
use Kolay\XlsxStream\Sinks\FileSink;
use Kolay\XlsxStream\Templates\Template;
use Kolay\XlsxStream\Writers\SinkableXlsxWriter;

class StateGuardsTest extends TestCase
{
    private string $testFile;

    /** @var list<string> */
    private array $templates = [];

    protected function setUp(): void
    {
        parent::setUp();
        $this->testFile = sys_get_temp_dir() . '/guards_test_' . uniqid() . '.xlsx';
    }

    protected function tearDown(): void
    {
        if (file_exists($this->testFile)) {
            unlink($this->testFile);
        }
        foreach ($this->templates as $path) {
            @unlink($path);
        }
        $this->templates = [];
        parent::tearDown();
    }

    /**
     * Entering template mode is a state transition like any other, and the
     * two ends of the writer's life are where it must be refused: a writer
     * that already started has emitted parts the template would replace, and
     * a closed one has nothing left to write.
     */
    public function test_template_mode_cannot_be_entered_after_start_or_after_close()
    {
        $template = Template::open($this->minimalTemplate());

        $started = new SinkableXlsxWriter(new FileSink($this->testFile));
        $started->startFile(['a']);
        try {
            $started->useTemplate($template);
            $this->fail('useTemplate() after startFile() should throw');
        } catch (XlsxStreamException $e) {
            $this->assertStringContainsString('already', $e->getMessage());
        }
        $started->writeRow(['x']);
        $started->finishFile();

        try {
            $started->useTemplate($template);
            $this->fail('useTemplate() after finishFile() should throw');
        } catch (XlsxStreamException $e) {
            $this->assertStringContainsString('closed', $e->getMessage());
        }

        $template->close();
    }

    public function test_a_finished_template_writer_refuses_further_work()
    {
        $out = sys_get_temp_dir().'/guards_tpl_'.uniqid().'.xlsx';
        $writer = SinkableXlsxWriter::fromTemplate(new FileSink($out), $this->minimalTemplate());
        $writer->sheet('Report', 2);
        $writer->writeRow(['x']);
        $writer->finishFile();

        foreach ([
            'sheet' => fn () => $writer->sheet('Report', 2),
            'writeRow' => fn () => $writer->writeRow(['y']),
            'finishFile' => fn () => $writer->finishFile(),
        ] as $operation => $call) {
            try {
                $call();
                $this->fail("{$operation}() after finishFile() should throw");
            } catch (XlsxStreamException $e) {
                $this->assertStringContainsString('closed', $e->getMessage(), $operation);
            }
        }

        @unlink($out);
    }

    /** A one-sheet layout with a header row and a single sample row. */
    private function minimalTemplate(): string
    {
        $path = sys_get_temp_dir().'/guards_tpl_src_'.uniqid().'.xlsx';
        $writer = SinkableXlsxWriter::createForFile($path);
        $writer->setHeaderStyle(['bold' => true]);
        $writer->startFile(['Value']);
        $writer->writeRow(['sample']);
        $writer->finishFile();
        $this->templates[] = $path;

        return $path;
    }

    public function test_write_row_before_start_throws()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('startFile() first');

        $writer->writeRow(['x']);
    }

    public function test_finish_file_before_start_throws()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('startFile() first');

        $writer->finishFile();
    }

    public function test_double_start_throws()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['A']);

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('already been started');

        $writer->startFile(['A']);
    }

    public function test_double_finish_throws_controlled_exception()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['A']);
        $writer->writeRow([1]);
        $writer->finishFile();

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('closed writer');

        $writer->finishFile();
    }

    public function test_write_after_finish_throws()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['A']);
        $writer->writeRow([1]);
        $writer->finishFile();

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('closed writer');

        $writer->writeRow([2]);
    }

    public function test_too_many_columns_in_headers_throws()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('exceeds Excel');

        $writer->startFile(array_fill(0, 17000, 'col'));
    }

    public function test_too_many_columns_in_row_throws()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(['A']);

        $this->expectException(XlsxStreamException::class);
        $this->expectExceptionMessage('exceeds Excel');

        $writer->writeRow(array_fill(0, 17000, 'x'));
    }

    public function test_max_columns_at_limit_succeeds()
    {
        $writer = new SinkableXlsxWriter(new FileSink($this->testFile));
        $writer->startFile(array_fill(0, 16384, 'h'));
        $writer->writeRow(array_fill(0, 16384, 'v'));
        $stats = $writer->finishFile();

        $this->assertEquals(1, $stats['rows']);
        $this->assertFileExists($this->testFile);
    }
}
