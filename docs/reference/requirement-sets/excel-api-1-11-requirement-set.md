---
title: Excel Требования к API JavaScript 1.11
description: Сведения о наборе требований ExcelApi 1.11.
ms.date: 04/01/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 1e347e0bb7b394676eccf422665c545b110b589d
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671438"
---
# <a name="whats-new-in-excel-javascript-api-111"></a>Новые возможности в Excel API JavaScript 1.11

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

В следующей таблице перечислены API Excel API JavaScript, за набором 1.11. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.11 или ранее, см. в Excel API в наборе требований [1.11](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)или ранее .

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureInfo)|Предоставляет сведения, основанные на текущих параметрах культуры системы.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalSeparator)|Получает строку, используемую в качестве десятичных сепараторов для числевых значений.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsSeparator)|Получает строку, используемую для отдельных групп цифр слева от десятичной для числимых значений.|
||[useSystemSeparators](/javascript/api/excel/excel.application#useSystemSeparators)|Указывает, включены ли системные Excel системы.|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|Получает объекты (например, люди), указанные в комментариях.|
||[richContent](/javascript/api/excel/excel.comment#richContent)|Получает богатое содержимое комментариев (например, упоминания в комментариях).|
||[разрешено](/javascript/api/excel/excel.comment#resolved)|Состояние потока комментариев.|
||[updateMentions(contentWithMentions: Excel. CommentRichContent)](/javascript/api/excel/excel.comment#updateMentions_contentWithMentions_)|Обновляет содержимое комментария с помощью специально отформатированной строки и списка упоминаний.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: CommentRichContent \| string, contentType?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#add_cellAddress__content__contentType_)|Создает новое примечание с указанным содержимым в определенной ячейке.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Адрес электронной почты объекта, упоминаемого в комментарии.|
||[id](/javascript/api/excel/excel.commentmention#id)|ID объекта.|
||[name](/javascript/api/excel/excel.commentmention#name)|Имя объекта, упомянутого в комментарии.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|Сущностям (например, людям), упомянутым в комментариях.|
||[разрешено](/javascript/api/excel/excel.commentreply#resolved)|Состояние ответа на комментарий.|
||[richContent](/javascript/api/excel/excel.commentreply#richContent)|Богатое содержимое комментариев (например, упоминания в комментариях).|
||[updateMentions(contentWithMentions: Excel. CommentRichContent)](/javascript/api/excel/excel.commentreply#updateMentions_contentWithMentions_)|Обновляет содержимое комментария с помощью специально отформатированной строки и списка упоминаний.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: CommentRichContent \| string, contentType?: Excel. ContentType)](/javascript/api/excel/excel.commentreplycollection#add_content__contentType_)|Создает ответ на комментарий для комментария.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|Массив, содержащий все сущностями (например, людьми), упомянутыми в комментарии.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richContent)|Указывает богатое содержимое комментария (например, комментарий контента с упоминаниями, первая упомянутая сущность имеет атрибут ID 0, а вторая упомянутая сущность имеет атрибут ID 1).|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Получает имя культуры в формате languagecode2-country/regioncode2 (например, "zh-cn" или "ru-ru").|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberFormat)|Определяет культурный формат отображения номеров.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberDecimalSeparator)|Получает строку, используемую в качестве десятичных сепараторов для числевых значений.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numberGroupSeparator)|Получает строку, используемую для отдельных групп цифр слева от десятичной для числимых значений.|
|[Range](/javascript/api/excel/excel.range)|[moveTo(destinationRange: Range \| string)](/javascript/api/excel/excel.range#moveTo_destinationRange_)|Перемещает значения ячейки, форматирование и формулы из текущего диапазона в диапазон назначения, заменяя старые сведения в этих ячейках.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent(amount: number)](/javascript/api/excel/excel.rangeformat#adjustIndent_amount_)|Регулирует отступ форматирования диапазона.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close_closeBehavior_)|Закрывает текущую книгу.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save_saveBehavior_)|Сохраняет текущую книгу.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onRowHiddenChanged)|Происходит, когда скрытое состояние одной или более строк изменилось на определенной таблице.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|Адрес диапазона, завершив вычисление.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onRowHiddenChanged)|Происходит, когда скрытое состояние одной или более строк изменилось на определенной таблице.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Получает адрес диапазона, представляющий измененную область конкретного листа.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changeType)|Получает тип изменений, которые представляют, как было вызвано событие.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetId)|Получает ID таблицы, в которой изменились данные.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
