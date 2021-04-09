---
title: Набор требований к API JavaScript Excel 1.11
description: Сведения о наборе требований ExcelApi 1.11.
ms.date: 04/01/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 7beabf94523164280d29c7f34c8b2c1003698bcc
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51650844"
---
# <a name="whats-new-in-excel-javascript-api-111"></a>Что нового в API JavaScript Excel 1.11

В ExcelApi 1.11 улучшена поддержка комментариев и элементов управления на уровне книг (например, сохранение и закрытие книги). Кроме того, добавлен доступ к настройкам культуры для учета локализации.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [Комментарии](../../excel/excel-add-ins-comments.md#mentions) |Теги и извества других пользователей книг с помощью комментариев. | [Комментарий](/javascript/api/excel/excel.comment), [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| Разрешение [комментариев](../../excel/excel-add-ins-comments.md#resolve-comment-threads) | Разрешить потоки комментариев и получить состояние разрешения. | [Comment](/javascript/api/excel/excel.comment) |
| [Параметры культуры](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Получает параметры культурной системы для книги, например форматирование номеров. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [Приложение NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [](/javascript/api/excel/excel.application) |
| [Вырезать и вклеить (moveTo)](../../excel/excel-add-ins-ranges-cut-copy-paste.md) | Реплицирует функции cut-and-paste в Excel для диапазона. | [Range](/javascript/api/excel/excel.range) |
| [Сохранение](../../excel/excel-add-ins-workbooks.md#save-the-workbook) и [закрытие](../../excel/excel-add-ins-workbooks.md#close-the-workbook) рабочей книги | Сохранение и закрытие книг. | [Workbook](/javascript/api/excel/excel.workbook) |
| События таблицы | Дополнительные события и сведения о событиях для вычислений таблиц и скрытых строк. | [WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs), [WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в API Excel JavaScript, за набором 1.11. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых требованиями API Excel JavaScript, установленных 1.11 или ранее, см. в справке к API Excel в наборе требований [1.11](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)или ранее .

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|Предоставляет сведения, основанные на текущих параметрах культуры системы.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|Получает строку, используемую в качестве десятичных сепараторов для числевых значений.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|Получает строку, используемую для отдельных групп цифр слева от десятичной для числимых значений.|
||[useSystemSeparators](/javascript/api/excel/excel.application#usesystemseparators)|Указывает, включены ли системные сепараторы Excel.|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|Получает объекты (например, люди), указанные в комментариях.|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|Получает богатое содержимое комментариев (например, упоминания в комментариях).|
||[разрешено](/javascript/api/excel/excel.comment#resolved)|Состояние потока комментариев.|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Обновляет содержимое комментария с помощью специально отформатированной строки и списка упоминаний.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Создает новое примечание с указанным содержимым в определенной ячейке.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Адрес электронной почты объекта, упоминаемого в комментарии.|
||[id](/javascript/api/excel/excel.commentmention#id)|ID объекта.|
||[name](/javascript/api/excel/excel.commentmention#name)|Имя объекта, упомянутого в комментарии.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|Сущностям (например, людям), упомянутым в комментариях.|
||[разрешено](/javascript/api/excel/excel.commentreply#resolved)|Состояние ответа на комментарий.|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|Богатое содержимое комментариев (например, упоминания в комментариях).|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Обновляет содержимое комментария с помощью специально отформатированной строки и списка упоминаний.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Создает ответ на примечание.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|Массив, содержащий все сущностями (например, людьми), упомянутыми в комментарии.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)|Указывает богатое содержимое комментария (например, комментарий содержимым с упоминаниями, первая упомянутая сущность имеет атрибут id 0, а вторая упомянутая сущность имеет атрибут id 1).|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Получает имя культуры в формате languagecode2-country/regioncode2 (например, "zh-cn" или "ru-ru").|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|Определяет культурный формат отображения номеров.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|Получает строку, используемую в качестве десятичных сепараторов для числевых значений.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|Получает строку, используемую для отдельных групп цифр слева от десятичной для числимых значений.|
|[Range](/javascript/api/excel/excel.range)|[moveTo(destinationRange: Range \| string)](/javascript/api/excel/excel.range#moveto-destinationrange-)|Перемещает значения ячейки, форматирование и формулы из текущего диапазона в диапазон назначения, заменяя старые сведения в этих ячейках.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent(amount: number)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Регулирует отступ форматирования диапазона.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Закрывает текущую книгу.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Сохраняет текущую книгу.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Происходит, когда скрытое состояние одной или более строк изменилось на определенной таблице.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|Адрес диапазона, завершив вычисление.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Происходит, когда скрытое состояние одной или более строк изменилось на определенной таблице.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Получает адрес диапазона, представляющий измененную область конкретного листа.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Получает тип изменений, которые представляют, как было вызвано событие.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Получает идентификатор листа, в котором изменены данные.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
