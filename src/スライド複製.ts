// スライド内の文字を抜き出す際
// グループ化されているテキストを抜き出すのが難しそうだった為
// グループ化を解除しています。
function slideinsert() {
  const mimeType = 'application/vnd.google-apps.presentation'; //フォルダ内から拾うデータの型をあらかじめ指定「GoogleSlide」
  const root = DriveApp.getFolderById(properties('staff_self')); //【2期生】自己紹介シート・登壇資料
  const copy_root = DriveApp.getFolderById(properties('master_slides')); //登壇資料マスタ
  // 【2期生】自己紹介シート・登壇資料からフォルダ一覧を取得
  const folders = root.getFolders();
  let file;
  // 取得したフォルダ一覧をループ処理
  // folders.hasNext()は、foldersの次のファイルがあるか判定。
  // なければfalseを返してループ終了
  while (folders.hasNext()) {
    const folder = folders.next();
    // スタッフフォルダ内の必要なデータを取得
    // _folder = 保存先
    // files = GoogleSlideのデータ一覧
    const _folder = folder.getFoldersByName('投影用').next();
    const files = folder.getFilesByType(mimeType);
    while (files.hasNext()) {
      // 取得したfilesの中から
      // ファイル名に「自己紹介シート」が含まれているファイルを選択
      file = files.next();
      if (file.getName().match(/自己紹介シート/)) {
        break;
      }
    }
    // 編集した、自己紹介シートを差し込みたい元データの一覧を取得
    // 一覧をループ処理して自己紹介シートを差し込んでいく。
    const copy_files = copy_root.getFilesByType(mimeType);
    while (copy_files.hasNext()) {
      const copy_file = copy_files.next();
      const file_name = copy_file.getName();
      const copy_id = copy_file.getId();
      const copy_slide = SlidesApp.openById(copy_id);
      // スライド一覧を取得
      const slides = copy_slide.getSlides();
      let check;
      for (let [index, slide] of slides.entries()) {
        const shapes = slide.getShapes();
        for (let shape of shapes) {
          const text = shape.getText().asRenderedString();
          if (text.includes('本日の内容')) {
            // スライドの中に、本日の内容が含まれているか判定。
            // 判定後、元データのファイル名を基準に
            // 自己紹介シートを差し込むスライド位置を調整。
            // スライドを差し込んだら、一度ファイルを開きなおしてから
            // createPDF関数を使用してPDFを出力。
            // 出力後は複製して自己紹介シートを挿入したGoogleSlideを削除する。
            if (file_name.match(/_Ver/)) {
            }
            else {
              index += 1;
            }
            const set_name = copy_file.getName().replace(/ のコピー/, '').replace(/【M】/, '');
            const make_copy = DriveApp.getFileById(copy_id).makeCopy(set_name, _folder);
            const id = make_copy.getId();
            const file_id = file.getId();
            const origin_slide = SlidesApp.openById(file_id).getSlides();
            const origin = origin_slide[origin_slide.length - 1];
            const presentation = SlidesApp.openById(id);
            presentation.insertSlide(index, origin);
            presentation.saveAndClose();
            createPDF(_folder.getId(), id, set_name);
            make_copy.setTrashed(true);
            check = true;
            break;
          }
          if (check) {
            break;
          }
        }
      }
    }
  }
}
const createPDF = (folder, id, file) => {
  DriveApp.getFolderById(folder).createFile(DriveApp.getFileById(id).getAs('application/pdf'))
    .setName(file);
};
