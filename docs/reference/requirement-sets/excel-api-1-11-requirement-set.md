---
title: Excel API JavaScript установлено 1.11
description: Сведения о наборе требований ExcelApi 1.11.
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
---

# <a name="whats-new-in-excel-javascript-api-111"></a>Новые возможности в Excel API JavaScript 1.11

В ExcelApi 1.11 улучшена поддержка комментариев и элементов управления на уровне книг (например, сохранение и закрытие книги). Кроме того, добавлен доступ к настройкам культуры для учета локализации.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [Комментарии](../../excel/excel-add-ins-comments.md#mentions) |Теги и извества других пользователей книг с помощью комментариев. | [Комментарий](/javascript/api/excel/excel.comment), [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| Разрешение [комментариев](../../excel/excel-add-ins-comments.md#resolve-comment-threads) | Разрешить потоки комментариев и получить состояние разрешения. | [Comment](/javascript/api/excel/excel.comment) |
| [Параметры культуры](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Получает параметры культурной системы для книги, например форматирование номеров. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [приложение NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [](/javascript/api/excel/excel.application) |
| [Вырезать и вклеить (moveTo)](../../excel/excel-add-ins-ranges-cut-copy-paste.md) | Реплицирует функции cut-and-paste в Excel для диапазона. | [Range](/javascript/api/excel/excel.range) |
| [Сохранение](../../excel/excel-add-ins-workbooks.md#save-the-workbook) и [закрытие](../../excel/excel-add-ins-workbooks.md#close-the-workbook) рабочей книги | Сохранение и закрытие книг. | [Workbook](/javascript/api/excel/excel.workbook) |
| События таблицы | Дополнительные события и сведения о событиях для вычислений таблиц и скрытых строк. | [WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs), [WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API Excel API JavaScript, за набором 1.11. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.11 или ранее, см. в Excel API в наборе требований [1.11 или ранее](/javascript/api/excel?view=excel-js-1.11&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#excel-excel-application-cultureinfo-member)|Предоставляет сведения, основанные на текущих параметрах культуры системы.|
||[decimalSeparator](/javascript/api/excel/excel.application#excel-excel-application-decimalseparator-member)|Получает строку, используемую в качестве десятичных сепараторов для числевых значений.|
||[thousandsSeparator](/javascript/api/excel/excel.application#excel-excel-application-thousandsseparator-member)|Получает строку, используемую для отдельных групп цифр слева от десятичной для числимых значений.|
||[useSystemSeparators](/javascript/api/excel/excel.application#excel-excel-application-usesystemseparators-member)|Указывает, включены ли системные Excel системы.|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#excel-excel-comment-mentions-member)|Получает объекты (например, люди), указанные в комментариях.|
||[разрешено](/javascript/api/excel/excel.comment#excel-excel-comment-resolved-member)|Состояние потока комментариев.|
||[richContent](/javascript/api/excel/excel.comment#excel-excel-comment-richcontent-member)|Получает богатое содержимое комментариев (например, упоминания в комментариях).|
||[updateMentions(contentWithMentions: Excel. CommentRichContent)](/javascript/api/excel/excel.comment#excel-excel-comment-updatementions-member(1))|Обновляет содержимое комментария с помощью специально отформатированной строки и списка упоминаний.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: CommentRichContent \| string, contentType?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-add-member(1))|Создает новое примечание с указанным содержимым в определенной ячейке.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#excel-excel-commentmention-email-member)|Адрес электронной почты объекта, упоминаемого в комментарии.|
||[id](/javascript/api/excel/excel.commentmention#excel-excel-commentmention-id-member)|ID объекта.|
||[name](/javascript/api/excel/excel.commentmention#excel-excel-commentmention-name-member)|Имя объекта, упомянутого в комментарии.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-mentions-member)|Сущностям (например, людям), упомянутым в комментариях.|
||[разрешено](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-resolved-member)|Состояние ответа на комментарий.|
||[richContent](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-richcontent-member)|Богатое содержимое комментариев (например, упоминания в комментариях).|
||[updateMentions(contentWithMentions: Excel. CommentRichContent)](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-updatementions-member(1))|Обновляет содержимое комментария с помощью специально отформатированной строки и списка упоминаний.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: CommentRichContent \| string, contentType?: Excel. ContentType)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-add-member(1))|Создает ответ на комментарий для комментария.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#excel-excel-commentrichcontent-mentions-member)|Массив, содержащий все сущностями (например, людьми), упомянутыми в комментарии.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#excel-excel-commentrichcontent-richcontent-member)|Указывает богатое содержимое комментария (например, комментарий контента с упоминаниями, первая упомянутая сущность имеет атрибут ID 0, а вторая упомянутая сущность имеет атрибут ID 1).|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#excel-excel-cultureinfo-name-member)|Получает имя культуры в формате languagecode2-country/regioncode2 (например, "zh-cn" или "ru-ru").|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#excel-excel-cultureinfo-numberformat-member)|Определяет культурный формат отображения номеров.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#excel-excel-numberformatinfo-numberdecimalseparator-member)|Получает строку, используемую в качестве десятичных сепараторов для числевых значений.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#excel-excel-numberformatinfo-numbergroupseparator-member)|Получает строку, используемую для отдельных групп цифр слева от десятичной для числимых значений.|
|[Range](/javascript/api/excel/excel.range)|[moveTo(destinationRange: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-moveto-member(1))|Перемещает значения ячейки, форматирование и формулы из текущего диапазона в диапазон назначения, заменяя старые сведения в этих ячейках.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent(amount: number)](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-adjustindent-member(1))|Регулирует отступ форматирования диапазона.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#excel-excel-workbook-close-member(1))|Закрывает текущую книгу.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#excel-excel-workbook-save-member(1))|Сохраняет текущую книгу.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowhiddenchanged-member)|Происходит, когда скрытое состояние одной или более строк изменилось на определенной таблице.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#excel-excel-worksheetcalculatedeventargs-address-member)|Адрес диапазона, завершив вычисление.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowhiddenchanged-member)|Происходит, когда скрытое состояние одной или более строк изменилось на определенной таблице.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-address-member)|Получает адрес диапазона, представляющий измененную область конкретного листа.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-changetype-member)|Получает тип изменений, которые представляют, как было вызвано событие.|
||[источник](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-source-member)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-worksheetid-member)|Получает ID таблицы, в которой изменились данные.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
