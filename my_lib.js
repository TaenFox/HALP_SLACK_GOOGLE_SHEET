function example() {
  var r = SpreadsheetApp.getActiveSpreadsheet();
  var s = r.getSheetByName('Реестр обращений');
  for (var j = 9064;j<9174;j++)
  {
    sp_message = lookUpTicketSP(execDescription(s,j,5))
    s.getRange(j,15).setValue(sp_message.ticket);
    s.getRange(j,16).setValue(sp_message.permalink);
    s.getRange(j,17).setValue(lookUpTicketHalp(sp_message.ticket));
  }

}

function execDescription(anySheet, rowSheet, colSheet)  //извлекает поле "описание" из содержания обращения
{
  var i = fix(anySheet.getRange(rowSheet, colSheet).getValue());
  var s;
  Logger.log(i);
  JSON.parse(i, function(key, value){if(key=="Описание"){s = value.slice(0,100)}});
  return s;
}

function fix(string)    //исправляем некорректные JSON
{
    let result = {}
    Array.from(
        string.matchAll(new RegExp('"([^"]+)":"(.*?)"[,}]', 'gsi')),
        match => result[match[1]] = match[2].replaceAll(/[\r\n"]/gim, ' ').replaceAll(/\s+/gim, ' ')
    )
    return JSON.stringify(result)
}


function lookUpTicketHalp (ticketNumber)  //возвращает только ссылку на сообщение
{
  var c = _slackSearch("halp_сортировка", "#"+ticketNumber, "Halp");
  if(c.ok==true)
  {
    if(c.messages.total>0 && c.messages.matches[0].permalink!=undefined)
      {
        var summary = c.messages.matches[0].permalink
      }else{
        Logger.log('Сообщение не найдено в канале halp сортировка')
      }
  }else{
    Logger.log("Ошибка отправки запроса: "+c.error)
  }
  return summary;
}

/*lookUpTicketSP ищет сообщение по тексту и возвращает массив данных по сообщению:
        error(ошибка обработки)
      или
        permalink(ссылка на сообщение),
        ticket(номер тикета),
        assign(кто взял в работу)
*/
function lookUpTicketSP(descriptionText)
{
  var c = _slackSearch("сопровождение_сервисов", descriptionText, "Halp");
  var summary; var permalink; var ticket; var assign;

  if(c.ok!=true)
  {
    Logger.log("Ошибка отправки запроса: "+c.error); return summary = {'error': true};
  };

  if(c.messages.total==0)
  {
    Logger.log('Сообщение не найдено в канале Сопровождение сервисов');return summary = {'error': true};
  };

  permalink = c.messages.matches[0].permalink;
  switch (_execStatusTicketFromText(c.messages.matches[0].text))
  {
    case 0: Logger.log("Произошла ошибка - это сообщение не тикет");return summary = {'error': true}; break;
    case 1: Logger.log("Закрытый тикет");ticket=_execNumberTicketFromText(c.messages.matches[0].text); break;
    default: Logger.log("Тикет в работе");ticket=_execNumberTicketFromText(c.messages.matches[0].text);break;
  };

  assign = c.messages.matches[0].attachments[0].footer;
  assign = assign.slice(13);
  summary = {
    'permalink': permalink,
    'ticket': ticket,
    'assign': assign
  };
  return summary;
}

function _execStatusTicketFromText (str)
{
  var s = str.match(/\[.*\]/gm);
  if (s == null) {return 0;}
  var j = s[0].slice(1, s[0].length-1);
  switch (j)
  {
    case "Закрыт": return 1; break;
    case "Передано разработчикам": return 2; break;
    case "Ожидаю ответа": return 3; break;
    case "Передано на выплату": return 4; break;
    default: return 9; break;
  }
}

function _execNumberTicketFromText (str)  //вытаскиваем из текста сообщения номер тикета
{
  var s = str.match(/\#[0-9]*/gm);
  if (s == null) {return 0;}
  return s[0].slice(1);
}

function _slackSearch(channelName, searchText, searchUser)  //выполняет метод поиска по каналу, автору и тексту (все аргументы обязательны)
{
  let query = `query=${encodeURIComponent('in:#'+channelName+' '+searchText+' from:@'+searchUser)}`;
  return _slackRequest("search.messages",query);
}

function _slackRequest(method, query)   //выполняет запросы к API Slack`а
{
  var token = "в это место нужно вставить токен пользователя";
  var auth = {
    'Authorization': 'Bearer '+ token
  }
  var options = {
    'headers': auth
  };
  let url = `https://slack.com/api/${method}?${query}`;
  var response = UrlFetchApp.fetch(url, options);
  var content = response.getContentText();
  return JSON.parse(content);
}
