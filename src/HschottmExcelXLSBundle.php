<?php

declare(strict_types=1);

namespace Hschottm\ExcelXLSBundle;

use Hschottm\ExcelXLSBundle\DependencyInjection\ExcelXLSExtension;
use Symfony\Component\HttpKernel\Bundle\Bundle;

class HschottmExcelXLSBundle extends Bundle
{
    public function getContainerExtension()
    {
        return new ExcelXLSExtension();
    }
}
