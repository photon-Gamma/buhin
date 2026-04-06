use serde::{Deserialize, Serialize};
use std::fs::File;
use std::io::BufReader;
use calamine::Data;
mod data_read;
use crate::data_read::data_read::bom_read;
use crate::data_read::data_read::param_read;

use rust_xlsxwriter::{
    Color,FormatBorder, Format, Workbook
};

// JSONの構造に合わせて構造体を定義
#[derive(Debug, Deserialize, Serialize)]
struct BomElement {
    name: String,
    figure_number: String,
    bom_file_path: String,
    bom_read_data: Vec<String>,
    param_file_path: String,
    param_read_data: Vec<String>,
    output_file_path: String,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // -----------------------------------
    // 部品表に追加する項目をJSONから読み込む
    // -----------------------------------
    // 1. ファイルを開く
    let file = File::open("./set_param.json")?;
    let reader = BufReader::new(file);
    // 2. JSONを構造体に変換（デシリアライズ）
    let user: BomElement = serde_json::from_reader(reader)?;

    // -----------------------------------
    // CADが出力したエクセルファイルを読み込み
    // -----------------------------------
    // 部品名のリストを用意, 検索に使う
    let part_name:String = "部材名".to_string();
    let mut part_name_data : Vec<Data>= Vec::new();
    // BOMデータを格納するベクターの初期化
    let mut data_vec : Vec<Vec<Data>> = Vec::new();
    // 完全一致の目的のヘッダー名をJSONから取得 (input)
    let target_header_v = user.bom_read_data;//所有権の移譲
    let bom_file_path = user.bom_file_path; //所有権の移譲
    bom_read(&mut data_vec, &mut part_name_data, &target_header_v, &bom_file_path, &part_name)?;
    
    // -----------------------------------
    // 追加したい要素のエクセルファイルの読み込み
    // -----------------------------------
    // データを格納するベクターの初期化
    let mut data_vec_plus : Vec<Vec<Data>> = Vec::new();
    // 完全一致の目的のヘッダー名をJSONから取得 (input)
    let target_header_v_plus = user.param_read_data;//所有権の移譲
    param_read(&mut data_vec_plus, &mut part_name_data, &target_header_v_plus, &user.param_file_path, &part_name)?;


    // -----------------------------------
    // エクセルファイルに書き込む
    // -----------------------------------
    // 1. ワークブックとワークシートの作成
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // 2. 書式（スタイル）の定義
    let _custom_font = Format::new()
        .set_font_name("BIZ UDGothic")    // フォント名（OSにインストールされているもの）
        .set_font_size(10.0)           // フォントサイズ（f64型）
        .set_font_color(Color::Black);   // フォントの色
    
        let _title_format = Format::new()
        .set_bold()
        .set_font_name("BIZ UDGothic")
        .set_font_size(15.0)
        .set_font_color(Color::Black)
        .set_background_color(Color::White)
        .set_border(FormatBorder::Thin);

    let _header_format = Format::new()
        .set_bold()
        .set_font_name("BIZ UDGothic")
        .set_font_size(13.0)
        .set_font_color(Color::Black)
        .set_background_color(Color::White)
        .set_border(FormatBorder::Thin);

    //let price_format = Format::new()
    //    .set_num_format("¥#,##0"); // 通貨フォーマット

    // 3. データの書き込み（書式を適用）
    //worksheet.merge_range(0, 0, 0, 3, &user.name, &_title_format)?;
    worksheet.write_with_format(0, 0, "製品名: ", &_title_format)?;
    worksheet.write_with_format(0, 1, &user.name, &_title_format)?;

    
    let columns_vec = [target_header_v, target_header_v_plus].concat();
    let data_list_vec = [data_vec, data_vec_plus].concat();
    
    for (col, value) in (columns_vec).iter().enumerate() {
        worksheet.write_with_format(1, col as u16, value.to_string(), &_header_format)?;
        // 列幅の調整, matchでいい感じに調整したい
        worksheet.set_column_width(col as u16, 20)?;
    }// 列のヘッダーを書き込むループ
    for (col, value_s) in data_list_vec.iter().enumerate() {
        
        for (row, value) in value_s.iter().enumerate() {
            // row は 0, 1, 2... と増えていくため、縦方向に書き込まれる
            worksheet.write_with_format((row+2) as u32, col as u16, value.to_string(), &_custom_font)?;
        }
    }// 部品情報を書き込むループ
    

    // fin. ワークブックの保存
    workbook.save(&user.output_file_path)?;
    
    Ok(())
}
