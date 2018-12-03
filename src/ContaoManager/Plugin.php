<?php

declare(strict_types=1);

namespace Hschottm\ExcelXLSBundle\ContaoManager;

use Contao\CoreBundle\ContaoCoreBundle;
use Contao\ManagerPlugin\Bundle\BundlePluginInterface;
use Contao\ManagerPlugin\Bundle\Config\BundleConfig;
use Contao\ManagerPlugin\Bundle\Parser\ParserInterface;
use Hschottm\ExcelXLSBundle\HschottmExcelXLSBundle;

/**
 * Plugin for the Contao Manager.
 *
 * @author Helmut Schottmüller (hschottm)
 */
class Plugin implements BundlePluginInterface
{
    /**
     * {@inheritdoc}
     */
    public function getBundles(ParserInterface $parser)
    {
        return [
             BundleConfig::create(HschottmExcelXLSBundle::class)
              ->setLoadAfter([ContaoCoreBundle::class])
              ->setReplace(['xls_export']),
         ];
    }
}
