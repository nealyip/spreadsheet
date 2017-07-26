<?php
/**
 * Created by PhpStorm.
 * User: neal.yip
 * Date: 26/7/2017
 * Time: 10:58
 */

namespace Nealyip\Spreadsheet;


use Illuminate\Support\ServiceProvider;

class SpreadsheetServiceProvider extends ServiceProvider
{
    /**
     * Bootstrap the application events.
     */
    public function boot()
    {

        $this->publishes(
            [
                __DIR__ . '/../config/spreadsheet.php' => config_path('spreadsheet.php'),
            ], 'config');

        $this->mergeConfigFrom(__DIR__ . '/../config/spreadsheet.php', 'spreadsheet');
    }

    /**
     * Register any application services.
     *
     * @return void
     */
    public function register()
    {

        $this->app->bind(Writer::class, function ($app) {
            return resolve('Nealyip\\Spreadsheet\\' . config('spreadsheet.writer') . 'Writer');
        });

        $this->app->bind(Reader::class, function ($app) {
            return resolve('Nealyip\\Spreadsheet\\' . config('spreadsheet.reader') . 'Reader');
        });

    }

}