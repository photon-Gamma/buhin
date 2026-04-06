


// -----------------------------------
// CADが出力したエクセルファイルを読み込み
// 2026/04/06
// -----------------------------------
//fn offset_func(v_offset_out: f32) -> (f32, f32, f32) 
pub mod data_read {
    use calamine::{Reader, open_workbook_auto, Data};
    pub fn bom_read(
            data_vec: &mut Vec<Vec<Data>>, 
            part_name_data: &mut Vec<Data>,//&x とすることで、所有権を奪わずに「借用」
            target_header_v: &Vec<String>,
            bom_file_path: &String,
            part_name: &String
        ) -> Result<(), Box<dyn std::error::Error>> {
        let mut workbook = open_workbook_auto(bom_file_path)?;
        // CADでシート名が変更されるため、.get(0) で最初のシート名を取得し、クローンして所有権を持つStringにする
        let first_sheet = workbook
            .sheet_names()
            .get(0)
            .ok_or("シートが見つかりません")?
            .clone();
        // worksheet_range は Result<Range<Data>, Error> を返すので ? で受ける
        let range = workbook.worksheet_range(&first_sheet)?;
        // 特定の列を抽出する
        // ヘッダー行（最初の行）からインデックスを特定する
        let mut rows = range.rows();
        let header_row = rows.next().ok_or("データが空です")?;
        for target_header in target_header_v {
            //println!("{}", target_header);
            let col_index = header_row//対象の列のインデックスを特定
                .iter()
                .position(|cell| cell.to_string() == *target_header)
                .ok_or(format!("ヘッダー '{}' が見つかりません", *target_header))?;
            // 特定したインデックスを使って、残りの行からデータを抽出
            let column_data: Vec<Data> = rows.clone()//rows自体もクローンして、所有権を持つデータに変換
                .filter_map(|row| row.get(col_index))// 指定した列が存在する場合のみ取得
                .cloned()// 所有権を持つデータに変換
                .collect();
            if *part_name == *target_header {
                //println!("部品名の列を見つけました: {}", *part_name);
                *part_name_data = column_data.clone(); // 部品名のデータを別のベクターに格納
            }
            data_vec.push(column_data);
            
        }//スコープを抜ける
        //println!("データベクター: {:?}", data_vec);
        //println!("part_name_data: {:?}", part_name_data);
        Ok(())
    }//BOM読み込みのスコープを抜ける
    pub fn param_read(
            data_vec_plus: &mut Vec<Vec<Data>>, 
            part_name_data: &mut Vec<Data>,//&x とすることで、所有権を奪わずに「借用」
            target_header_v_plus: &Vec<String>,
            param_file_path: &String,
            part_name: &String
        ) -> Result<(), Box<dyn std::error::Error>> {
        let mut workbook = open_workbook_auto(param_file_path)?;
        let sheet_name = "Sheet1";
        let range = workbook
            .worksheet_range(sheet_name)?;
        let mut rows = range.rows();
        let header_row = rows.next().ok_or("データが空です")?;
        
        // 部品リストの部材名と一致する要素を抽出
        let col_index = header_row//対象の列のインデックスを特定
                .iter()
                .position(|cell| cell.to_string() == *part_name)
                .ok_or(format!("ヘッダー '{}' が見つかりません", *part_name))?;
        let read_data_i: Vec<Data> = rows.clone()//rows自体もクローンして、所有権を持つデータに変換
                .filter_map(|row| row.get(col_index))// 指定した列が存在する場合のみ取得
                .cloned()// 所有権を持つデータに変換
                .collect();
        let matches: Vec<usize> = read_data_i.iter()
            .enumerate()
            .filter(|(_, item)| part_name_data.contains(item)) // part_name_dataに含まれているか
            .map(|(i, _)| i)
            .collect();
        //println!("matches: {:?}", matches);

        for target_header in target_header_v_plus {//&x とすることで、所有権を奪わずに「借用」してループを回す
            //println!("{}", target_header);
            let col_index = header_row//対象の列のインデックスを特定
                .iter()
                .position(|cell| cell.to_string() == *target_header)
                .ok_or(format!("ヘッダー '{}' が見つかりません", *target_header))?;
            // 特定したインデックスを使って、残りの行からデータを抽出
            let column_data: Vec<Data> = rows.clone()//rows自体もクローンして、所有権を持つデータに変換
                .filter_map(|row| row.get(col_index))// 指定した列が存在する場合のみ取得
                .cloned()// 所有権を持つデータに変換
                .collect();
            let result: Vec<Data> = matches.iter()
                .map(|&i| column_data[i].clone())
                .collect();
            data_vec_plus.push(result);
            
        }//スコープを抜ける
        //println!("データベクター: {:?}", data_vec_plus);
        //println!("データベクター: {:?}", data_vec_plus.get(0).and_then(|row| row.get(0)));
        Ok(())
    }
}