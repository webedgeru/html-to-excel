<?php

namespace yii2\htmltoexcel;

use yii\base\Application;
use yii\base\BootstrapInterface;

/**
 * Bootstrap class for yii2-html-to-excel extension.
 */
class Bootstrap implements BootstrapInterface
{
    /**
     * @inheritdoc
     */
    public function bootstrap($app)
    {
        // Set alias for the extension
        if (!\Yii::getAlias('@yii2/htmltoexcel', false)) {
            \Yii::setAlias('@yii2/htmltoexcel', __DIR__);
        }
    }
}
