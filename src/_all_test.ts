const difftes = () => {
  const obj = new LogclockCheck()
  console.log(Object.keys(obj))
}
const createMap = (values, label, keys) => {
  const obj = {}
  keys.forEach(key => {
    const type = Object.prototype.toString.call(key)
    if (type == '[object Array]') {
      obj['member'] = values.filter((value, index) => index >= key[0] && index <= key[key.length - 1])
    }
  })
}
const resultCreate = (values, label, keys) => {
  const obj = {}
  keys.forEach(key => {

  })
}
const checkCreate = (values, label, keys) => {
  const obj = {}
  keys.forEach(key => {
    const value = values[label.indexOf(key)]
    switch (key) {
      case '日程': obj['day'] = new Times(value).MMdd()
        break
      case '会場\n名称': obj['venue'] = value
        break
      case '開催No.': obj['serial'] = String(value)
    }
  })
  return obj
}
class LogclockValues {
  constructor(values, label) {
    const Check = (bool, origin) => {
      if (!bool) { return origin }
      return bool
    }
    const flag = values[label.indexOf('開催\n可否')] == '中止'
    return ['日程', '会場\n名称', '開始', 'メイン\n講師', 'サポート講師', 'サポート2', 'サポート3', 'サポート4', 'サポート5', 'Check1', 'Check2', 'Check3']
      .map(key => {
        const value = values[label.indexOf(key)]
        switch (key) {
          case '日程': return new Times(value).MMdd()
          case '開始': return new Times(value).Hmm()
          case 'Check1': return Check(value, values[label.indexOf('現場登録')])
          case 'Check2': return Check(value, values[label.indexOf('お仕事スケジュール')])
          case 'Check3':
            if (flag) {

            }
            return Check(value, values[label.indexOf('キャスティング')])
          default: return value
        }
      })
  }
}
// class VenueCallSheet {
//   // origin_label
//   // spread

//   constructor() {
//     this.spread = mainData_('vc')
//     this.origin_label = JSON.parse(properties('sheet_label'))
//   }
//   labelCreate(values) {
//     //新会場連絡シートのラベルを設定。
//     //アサインシート側のラベルに名称を変更。
//     return values.map(value => value
//       .replace(/(^会場$|^会場名$)/, '会場\n名称')
//       .replace(/^可否$/, '開催\n可否')
//       .replace(/^日付$/, '日程')
//       .replace(/(^メイン$|^変更後メンバー$)/, 'メイン\n講師')
//       .replace(/(^サポート1$|^メンバー1$)/, 'サポート講師')
//       .replace(/^メンバー/, 'サポート')
//       .replace(/^会場特性$/, '会場運用上\n注意点')
//       .replace(/^備考$/, 'カリキュラム\n補足')
//       .replace(/^予定$/, '参加予定人数'));
//   }
// }
// class Values extends VenueCallSheet {
//   /**
//    * @_label ラベルデータ:string[]
//    * @values 元データ
//    * @rerutn ラベルデータに基づいて、元データを整形して配列として返す。
//    */
//   constructor(_label, values) {
//     super()
//     const flag = (values[_label.indexOf('講師')] == 'エムジー')
//     const start = values[_label.indexOf('開始')]
//     const finish = values[_label.indexOf('終了')]
//     return this.label.flatMap(key => {
//       const value = values[_label.indexOf(key)]
//       switch (key) {
//         case '開催No.':
//         case '会場\n番号':
//         case '主催者TEL':
//         case '会場TEL':
//         case '通し番号': return String(value)
//         case '日程':
//         case '更新日':
//         case '確認日': return new Times(value).MMdd()
//         case '開始':
//         case '終了': return new Times(value).Hmm()
//         case '集合': return new Times(start).meetingTime(flag)
//         case '解散': return new Times(finish).endTime()
//         case 'LOG住所': return new AddressWork(value).onlyAddress()
//         case 'LOG建物': return new AddressWork(value).building()
//         case 'LOG主催者TEL':
//         case 'LOG会場TEL': return new TextNumbers(value).onlyPhoneNumber()
//         default:
//           if (_label.indexOf(key) == -1) {
//             return ''
//           }
//           return value
//       }
//     })
//   }
// }
// class VenueCall extends VenueCallSheet {
//   constructor() {
//     super()
//     const sheet = this.spread.getSheetByName('会場連絡')

//     this.sheet = sheet
//   }
//   checkData() { }
//   memoData() { }
// }

// class CastingSheet extends VenueCallSheet {
//   constructor() {
//     super()
//     const sheet = this.spread.getSheetByName('キャスティング')

//     this.sheet = sheet
//   }

// }
// class Numbers {
//   serial
//   ven_num
//   organizer_phone
//   venue_phone
//   log_organizer_phone
//   log_venue_phone
//   limit
//   plan_to_people
//   carry
//   assign_people
//   set_number
// }
// class Times {
//   day
//   start
//   finish
//   meeting
//   leave

// }
// class Strings {
//   area
//   hold
//   venue_name
//   venue_address
//   log_address
//   log_building
//   corse
//   main
//   support
//   store
//   sad_support
//   caution
//   sb_manager
//   venue_manager
// }

// class Booleans {
//   sad_flag
//   main_flag
// }