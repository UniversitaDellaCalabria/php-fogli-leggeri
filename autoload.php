<?php

declare(strict_types=1);

/**
 * Per chi non usa Composer: `require '/percorso/php-fogli-leggeri/autoload.php';` e le classi del pacchetto
 * (namespace Dodaroe\FogliLeggeri) si caricano da sole. Con Composer questo file non serve.
 */
spl_autoload_register(static function (string $classe): void {
    $prefisso = 'Dodaroe\\FogliLeggeri\\';
    if (!str_starts_with($classe, $prefisso)) {
        return;
    }
    $file = __DIR__ . '/src/' . str_replace('\\', '/', substr($classe, strlen($prefisso))) . '.php';
    if (is_file($file)) {
        require $file;
    }
});
