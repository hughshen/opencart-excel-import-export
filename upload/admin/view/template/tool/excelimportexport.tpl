<?php echo $header; ?><?php echo $column_left; ?>
<div id="content">
  <div class="page-header">
    <div class="container-fluid">
      <h1><?php echo $heading_title; ?></h1>
      <ul class="breadcrumb">
        <?php foreach ($breadcrumbs as $breadcrumb) { ?>
        <li><a href="<?php echo $breadcrumb['href']; ?>"><?php echo $breadcrumb['text']; ?></a></li>
        <?php } ?>
      </ul>
    </div>
  </div>

  <!-- Product import -->
  <div class="container-fluid">
    <?php if ($error_warning) { ?>
    <div class="alert alert-danger"><i class="fa fa-exclamation-circle"></i> <?php echo $error_warning; ?>
      <button type="button" class="close" data-dismiss="alert">&times;</button>
    </div>
    <?php } ?>
    <?php if ($success) { ?>
    <div class="alert alert-success"><i class="fa fa-check-circle"></i> <?php echo $success; ?>
      <button type="button" form="form-backup" class="close" data-dismiss="alert">&times;</button>
    </div>
    <?php } ?>
    <div class="panel panel-default">
      <div class="panel-heading">
        <h3 class="panel-title"><i class="fa fa-exchange"></i> <?php echo $product_import_heading_title; ?></h3>
      </div>
      <div class="panel-body">
        <form action="<?php echo $import_products; ?>" method="post" enctype="multipart/form-data" id="products-import-form" class="form-horizontal">
          <div class="form-group">
            <label class="col-sm-2 control-label"><?php echo $text_choose_import; ?></label>
            <div class="col-sm-10">
              <input type="file" name="upload" id="upload" />
            </div>
          </div>
          <div class="form-group">
            <label class="col-sm-2 control-label">&nbsp;</label>
            <div class="col-sm-1">
              <a class="btn btn-primary" onclick="importProducts();"><span><?php echo $button_import; ?></span></a>
            </div>
          </div>
          <div class="row">
            <label class="col-sm-2 control-label"><?php echo $button_export; ?></label>
            <div class="col-sm-2">
              <a class="btn btn-primary" href="<?php echo $export_sample; ?>"><?php echo $text_export_sample; ?></span></a>
            </div>
            <div class="col-sm-2">
              <a class="btn btn-primary" href="<?php echo $export_id_field; ?>"><?php echo $text_export_id_field; ?></span></a>
            </div>
          </div>      
        </form>
      </div>
    </div>
  </div>

  <!-- Orders location export -->
  <div class="container-fluid">
    <div class="panel panel-default">
      <div class="panel-heading">
        <h3 class="panel-title"><i class="fa fa-exchange"></i> <?php echo $orders_location_heading_title; ?></h3>
      </div>
      <div class="panel-body">
        <form action="<?php echo $export_orders_location; ?>" method="post" enctype="multipart/form-data" id="orders-location-form" class="form-horizontal">
          <div class="table-responsive">
            <table id="discount" class="table table-striped table-bordered table-hover">
              <thead>
                <tr>
                  <td class="text-left"><?php echo $text_start_date; ?></td>
                  <td class="text-left"><?php echo $text_end_date; ?></td>
                  <td></td>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td class="text-left">
                      <div class="input-group date">
                          <input type="text" name="start_date" value="" placeholder="<?php echo $text_start_date; ?>" data-date-format="YYYY-MM-DD" class="form-control">
                          <span class="input-group-btn"><button type="button" class="btn btn-default"><i class="fa fa-calendar"></i></button></span>
                      </div>
                  </td>
                  <td class="text-left">
                      <div class="input-group date">
                          <input type="text" name="end_date" value="" placeholder="<?php echo $text_end_date; ?>" data-date-format="YYYY-MM-DD" class="form-control">
                          <span class="input-group-btn"><button type="button" class="btn btn-default"><i class="fa fa-calendar"></i></button></span>
                      </div>
                  </td>
                  <td class="text-left">
                    <a class="btn btn-primary" onclick="ordersLocationExport()"><span><?php echo $button_export; ?></span></a>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>       
        </form>
      </div>
    </div>
  </div>

  <!-- Sales order export -->
  <div class="container-fluid">
    <div class="panel panel-default">
      <div class="panel-heading">
        <h3 class="panel-title"><i class="fa fa-exchange"></i> <?php echo $sales_export_heading_title; ?></h3>
      </div>
      <div class="panel-body">
        <form action="<?php echo $export_sales_order; ?>" method="post" enctype="multipart/form-data" id="sales-order-form" class="form-horizontal">
          <div class="table-responsive">
            <table id="discount" class="table table-striped table-bordered table-hover">
              <thead>
                <tr>
                  <td class="text-left"><?php echo $text_start_date; ?></td>
                  <td class="text-left"><?php echo $text_end_date; ?></td>
                  <td></td>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td class="text-left">
                      <div class="input-group date">
                          <input type="text" name="start_date" value="" placeholder="<?php echo $text_start_date; ?>" data-date-format="YYYY-MM-DD" class="form-control">
                          <span class="input-group-btn"><button type="button" class="btn btn-default"><i class="fa fa-calendar"></i></button></span>
                      </div>
                  </td>
                  <td class="text-left">
                      <div class="input-group date">
                          <input type="text" name="end_date" value="" placeholder="<?php echo $text_end_date; ?>" data-date-format="YYYY-MM-DD" class="form-control">
                          <span class="input-group-btn"><button type="button" class="btn btn-default"><i class="fa fa-calendar"></i></button></span>
                      </div>
                  </td>
                  <td class="text-left">
                    <a class="btn btn-primary" onclick="salesOrderExport()"><span><?php echo $button_export; ?></span></a>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>       
        </form>
      </div>
    </div>
  </div>
</div>

<script type="text/javascript"><!--
$("input[value=pid]").attr("checked",true)
$(".page").hide();
$(".pid").show();

$(function() {
    $("input[value=page]").click(function() {
        $(".pid").hide();
        $(".page").show(0);
    });
    $("input[value=pid]").click(function() {
        $(".page").hide();
        $(".pid").show();
    });

    $('.date').datetimepicker({
      pickTime: false
    });
});

function checkFileSize(id) {
  // See also http://stackoverflow.com/questions/3717793/javascript-file-upload-size-validation for details
  var input, file, file_size;

  if (!window.FileReader) {
    // The file API isn't yet supported on user's browser
    return true;
  }

  input = document.getElementById(id);
  if (!input) {
    // couldn't find the file input element
    return true;
  }
  else if (!input.files) {
    // browser doesn't seem to support the `files` property of file inputs
    return true;
  }
  else if (!input.files[0]) {
    // no file has been selected for the upload
    alert( "<?php echo $error_select_file; ?>" );
    return false;
  }
  else {
    file = input.files[0];
    file_size = file.size;
    <?php if (!empty($post_max_size)) { ?>
    // check against PHP's post_max_size
    post_max_size = <?php echo $post_max_size; ?>;
    if (file_size > post_max_size) {
      alert( "<?php echo $error_post_max_size; ?>" );
      return false;
    }
    <?php } ?>
    <?php if (!empty($upload_max_filesize)) { ?>
    // check against PHP's upload_max_filesize
    upload_max_filesize = <?php echo $upload_max_filesize; ?>;
    if (file_size > upload_max_filesize) {
      alert( "<?php echo $error_upload_max_filesize; ?>" );
      return false;
    }
    <?php } ?>
    return true;
  }
}

function importProducts() {
  if (checkFileSize('upload')) {
    $('#products-import-form').submit();
  }
}

function ordersLocationExport() {
  $('#orders-location-form').submit();
}

function salesOrderExport() {
  $('#sales-order-form').submit();
}
</script>

<?php echo $footer; ?>