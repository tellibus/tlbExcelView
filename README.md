tlbExcelView
============

Another Yii CGridView-to-Excel exporter using [PHPExcel](http://phpexcel.codeplex.com/), based on [EExcelView Yii extension](http://yiiframework.com/extension/eexcelview/)

Tested with:

 - Yii 1.1.10
 - PHPExcel 1.7.7


Homepage
--------
[http://tellibus.com/lab/tlbExcelView](http://tellibus.com/lab/tlbExcelView)

Installation
------------

 - Create a folder named **phpexcel** in your **/protected/extensions** folder
 - Download PHPExcel from [http://phpexcel.codeplex.com/](http://phpexcel.codeplex.com/)
 - Unpack its **Classes** folder in your **/protected/extensions/phpexcel**
 - Copy the **tlbExcelView.php** file in your widgets directory, in this example **/protected/components/widgets**

Features
--------

 - Automatic formatting of header, body and footer
 - Automatic formatting of numbers
 - Automatic sum in the footer
 - Automatic page formatting (with page header and footer, automatic print area…)
 - Nearly all properties can be overridden.


Example of use
--------------

This is an example of use of tlbExcelView in the controller and view:

### Controller

Based on the standard Gii / Giix admin action

```php
<?php public function actionAdmin() {
    $model = new Model('search');
    $model->unsetAttributes();
    if (isset($_GET['Model'])) {
        $model->attributes = $_GET['Model'];
    }
    if (isset($_GET['export'])) {
        $production = 'export';
    } else {
        $production = 'grid';
    }
    $this->render('admin', array('model' => $model, 'production' => $production));
} ?>
```

### "admin" view

```php
<?php Yii::app()->clientScript->registerScript('search', "
    $('#exportToExcel').click(function(){
        window.location = '". $this->createUrl('admin')  . "?' + $(this).parents('form').serialize() + '&export=true';
        return false;
    });
    $('.search-form form').submit(function(){
        $.fn.yiiGridView.update('some-grid', {
            data: $(this).serialize()
        });
        return false;
    });
"); ?>
…
<div class="search-form" style="display:block">
<?php $this->renderPartial('_search', array('model' => $model)); ?>
</div><!-- search-form -->

…
<?php $this->widget('application.components.widgets.tlbExcelView', array(
    'id'                   => 'some-grid',
    'dataProvider'         => $model->search(),
    'grid_mode'            => $production, // Same usage as EExcelView v0.33
    //'template'           => "{summary}\n{items}\n{exportbuttons}\n{pager}",
    'title'                => 'Some title - ' . date('d-m-Y - H-i-s'),
    'creator'              => 'Your Name',
    'subject'              => mb_convert_encoding('Something important with a date in French: ' . utf8_encode(strftime('%e %B %Y')), 'ISO-8859-1', 'UTF-8'),
    'description'          => mb_convert_encoding('Etat de production généré à la demande par l\'administrateur (some text in French).', 'ISO-8859-1', 'UTF-8'),
    'lastModifiedBy'       => 'Some Name',
    'sheetTitle'           => 'Report on ' . date('m-d-Y H-i'),
    'keywords'             => '',
    'category'             => '',
    'landscapeDisplay'     => true, // Default: false
    'A4'                   => true, // Default: false - ie : Letter (PHPExcel default)
    'pageFooterText'       => '&RThis is page no. &P of &N pages', // Default: '&RPage &P of &N'
    'automaticSum'         => true, // Default: false
    'decimalSeparator'     => ',', // Default: '.'
    'thousandsSeparator'   => '.', // Default: ','
    //'displayZeros'       => false,
    //'zeroPlaceholder'    => '-',
    'sumLabel'             => 'Column totals:', // Default: 'Totals'
    'borderColor'          => '00FF00', // Default: '000000'
    'bgColor'              => 'FFFF00', // Default: 'FFFFFF'
    'textColor'            => 'FF0000', // Default: '000000'
    'rowHeight'            => 45, // Default: 15
    'headerBorderColor'    => 'FF0000', // Default: '000000'
    'headerBgColor'        => 'CCCCCC', // Default: 'CCCCCC'
    'headerTextColor'      => '0000FF', // Default: '000000'
    'headerHeight'         => 10, // Default: 20
    'footerBorderColor'    => '0000FF', // Default: '000000'
    'footerBgColor'        => '00FFCC', // Default: 'FFFFCC'
    'footerTextColor'      => 'FF00FF', // Default: '0000FF'
    'footerHeight'         => 50, // Default: 20
    'columns'              => $grid // an array of your CGridColumns
)); ?>
```

### "_search" view
```php
<?php $form = $this->beginWidget('GxActiveForm', array(
    'action' => Yii::app()->createUrl($this->route),
    'method' => 'get',
)); ?>
…
    <div class="row buttons">
        <?php echo GxHtml::submitButton(Yii::t('app', 'Search')); ?>
        <?php echo GxHtml::button(Yii::t('app', 'Export to Excel (xls)'), array('id' => 'exportToExcel')); ?>
    </div>
<?php $this->endWidget(); ?>
```