<?php
/*
Plugin Name: Export factfind to Excel
Description: Export custom post type "factfind" single post data to Excel format
*/

function export_links_to_excel() {
  if (isset($_GET['export_links'])) {
      $args = array(
          'post_type' => 'factfind',
          'posts_per_page' => -1
      );
      $query = new WP_Query($args);
      $filename = 'factfind-' . date('YmdHis') . '.xlsx';

      // Load PHPExcel library
      require_once(plugin_dir_path(__FILE__) . 'PHPExcel/PHPExcel.php');

      // Create new PHPExcel object
      $objPHPExcel = new PHPExcel();

      // Set properties of Excel file
      $objPHPExcel->getProperties()->setCreator('Your Name')
          ->setLastModifiedBy('Your Name')
          ->setTitle('Links Data Export')
          ->setSubject('Links Data')
          ->setDescription('Links data export in Excel format')
          ->setKeywords('links data excel export')
          ->setCategory('Links Data Export');

      // Add data to Excel file
      $objPHPExcel->setActiveSheetIndex(0);
      $objPHPExcel->getActiveSheet()->setTitle('Links Data');

      // Add column headers
      $objPHPExcel->getActiveSheet()->setCellValue('A1', 'Title');
      $col = 'B';
      $fields = array();
      while($query->have_posts()) {
          $query->the_post();
          $postId = get_the_ID();
          $postTitle = get_the_title();
          $fields[$postId] = get_post_custom($postId);
          foreach($fields[$postId] as $key => $value) {
              $objPHPExcel->getActiveSheet()->setCellValue($col . '1', $key);
              $col++;
          }
      }

      // Add row data
      $row = 2;
      while($query->have_posts()) {
          $query->the_post();
          $postId = get_the_ID();
          $postTitle = get_the_title();
          $objPHPExcel->getActiveSheet()->setCellValue('A' . $row, $postTitle);
          $col = 'B';
          foreach($fields[$postId] as $key => $value) {
              if(count($value) == 1) {
                  $objPHPExcel->getActiveSheet()->setCellValue($col . $row, $value[0]);
              } else {
                  $objPHPExcel->getActiveSheet()->setCellValue($col . $row, implode(',', $value));
              }
              $col++;
          }
          $row++;
      }

      // Save Excel file
      header('Content-Type: application/vnd.ms-excel');
      header('Content-Disposition: attachment;filename="' . $filename . '"');
      header('Cache-Control: max-age=0');
      $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
      $objWriter->save('php://output');
      exit;
  }
}

add_action('admin_init', 'export_factfind_to_excel');

function add_export_factfind_link($actions, $post) {
  if ($post->post_type == 'factfind') {
      $url = wp_nonce_url(admin_url('admin.php?action=export_factfind&post=' . $post->ID), 'export_links');
      $actions['export_factfind'] = '<a href="' . $url . '">Export to Excel</a>';
  }
  return $actions;
}

add_filter('post_row_actions', 'add_export_factfind_link', 10, 2);