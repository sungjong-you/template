/**
 * シートデータを管理するためのクラス
 */
class ShDataObj {
  /**
   * @param {Object} sh - Google Sheets のシートオブジェクト
   * @param {number|string} [last_sh_row_num=""] - 処理する最後の行番号。0 以上の数字を指定した場合、その行までのデータを取得する。空文字の場合は全体のデータを取得。
   * @param {any} [parm=null] - 任意の追加パラメータ。文字列、オブジェクト、その他のデータを含めることができる。
   */


  constructor(sh, last_sh_row_num = "") {
    this.sh = sh;
    this.arys = (last_sh_row_num !== "" && last_sh_row_num > 0)
      ? this.sh.getRange(1, 1, last_sh_row_num, this.sh.getLastColumn()).getValues()
      : this.sh.getDataRange().getValues();

    this.row_obj = (() => {
      let obj = {};
      for (let i = 0; i < this.arys.length; i++) {
        if (this.arys[i][0] !== "" && this.arys[i][0] !== null) obj[this.arys[i][0]] = i + 1;
      }
      return obj;
    })();

    this.col_obj = (() => {
      let obj = {};
      for (let i = 0; i < this.arys[0].length; i++) {
        if (this.arys[0][i] !== "" && this.arys[0][i] !== null) obj[this.arys[0][i]] = i + 1;
      }
      return obj;
    })();
  }


  /**
     * 任意の列の空白ではない最終行番号を取得(obj生成時にlast_sh_row_numの指定がある場合はその値以内に限定される)
     * @param {number[]} sh_col_num_ary - 列番号の配列（1から始まる列番号）
     * @returns {number} 指定された列の中で最終行番号
     */
  getLastNonEmptyShRowInColumns(sh_col_num_ary) {
    let lastRow = 0;
    sh_col_num_ary.forEach(sh_col_num => {
      for (let rowNum = this.arys.length - 1; rowNum >= 0; rowNum--) {
        if (this.arys[rowNum][sh_col_num - 1] !== "" && this.arys[rowNum][sh_col_num - 1] !== null) {
          lastRow = Math.max(lastRow, rowNum + 1);
          break;
        }
      }
    });
    return lastRow;
  }

}




