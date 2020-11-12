---
title: Набор обязательных элементов API JavaScript для Excel 1,11
description: Сведения о наборе требований ExcelApi 1,11.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 2ca378364cacab58000607aef15f450e14545ac6
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996272"
---
# <a name="whats-new-in-excel-javascript-api-111"></a>Новые возможности API JavaScript для Excel 1,11

В ExcelApi 1,11 улучшена поддержка комментариев и элементов управления уровня книги (например, при сохранении и закрытии книги). Кроме того, добавлен доступ к параметрам культуры для помощи в учетной записи для локализации.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [Упоминание](../../excel/excel-add-ins-comments.md#mentions) комментариев |Теги и уведомляет других пользователей книги с помощью комментариев. | [Comment](/javascript/api/excel/excel.comment), [комментричконтент](/javascript/api/excel/excel.commentrichcontent) |
| [Разрешение](../../excel/excel-add-ins-comments.md#resolve-comment-threads) комментариев | Разрешение потоков комментариев и получение состояния разрешения. | [Comment](/javascript/api/excel/excel.comment) |
| [Параметры культуры](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Получает региональные параметры системы для книги, например форматирование чисел. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [Application](/javascript/api/excel/excel.application) |
| [Вырезать и вставить (moveTo)](../../excel/excel-add-ins-ranges-advanced.md#cut-copy-and-paste) | Реплицирует функции вырезания и вставки в Excel для диапазона. | [Range](/javascript/api/excel/excel.range) |
| [Сохранение](../../excel/excel-add-ins-workbooks.md#save-the-workbook) и [закрытие](../../excel/excel-add-ins-workbooks.md#close-the-workbook) рабочей книги | Сохранение и закрытие книг. | [Workbook](/javascript/api/excel/excel.workbook) |
| События листа | Дополнительные события и сведения о событиях для вычислений и скрытых строк в листах. | [Воркшиткалкулатедевентаргс](/javascript/api/excel/excel.worksheetcalculatedeventargs), [воркшитровхидденчанжедевентаргс](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в наборе обязательных элементов API JavaScript для Excel 1,11. Чтобы просмотреть справочную документацию по API для всех API, поддерживаемых набором обязательных элементов API JavaScript для Excel 1,11 или более ранней версии, обратитесь к разделам [API Excel в наборе требований 1,11](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|Предоставляет сведения, основанные на текущих параметрах языковых параметров системы.|
||[деЦималсепаратор](/javascript/api/excel/excel.application#decimalseparator)|Получает строку, используемую в качестве десятичного разделителя для числовых значений.|
||[саусандссепаратор](/javascript/api/excel/excel.application#thousandsseparator)|Получает строку, используемую для разделения групп цифр слева от десятичного разделителя для числовых значений.|
||[усесистемсепараторс](/javascript/api/excel/excel.application#usesystemseparators)|Указывает, включены ли системные разделители Excel.|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|Получает объекты (например, людей), которые упоминаются в комментариях.|
||[ричконтент](/javascript/api/excel/excel.comment#richcontent)|Получает содержимое форматированного комментария (например, упоминание в комментариях).|
||[определяем](/javascript/api/excel/excel.comment#resolved)|Состояние цепочки комментариев.|
||[Упдатементионс (Контентвисментионс: Excel. Комментричконтент)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Обновляет содержимое комментария с помощью специально отформатированной строки и списка упоминаний.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[Add (Целладдресс: \| строка Range, Content: комментричконтент \| String, ContentType?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Создает новое примечание с указанным содержимым в определенной ячейке.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Адрес электронной почты объекта, который упоминается в примечании.|
||[id](/javascript/api/excel/excel.commentmention#id)|Идентификатор объекта.|
||[name](/javascript/api/excel/excel.commentmention#name)|Имя объекта, который упоминается в примечании.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|Сущности (например, люди), которые упоминаются в комментариях.|
||[определяем](/javascript/api/excel/excel.commentreply#resolved)|Состояние ответа на комментарий.|
||[ричконтент](/javascript/api/excel/excel.commentreply#richcontent)|Содержимое форматированного комментария (например, упоминание в комментариях).|
||[Упдатементионс (Контентвисментионс: Excel. Комментричконтент)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Обновляет содержимое комментария с помощью специально отформатированной строки и списка упоминаний.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[Добавить (контент: \| строка комментричконтент, ContentType?: Excel. ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Создает ответ на примечание.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|Массив, содержащий все сущности (например, люди), упомянутые в комментарии.|
||[ричконтент](/javascript/api/excel/excel.commentrichcontent#richcontent)|Задает расширенное содержимое комментария (например, закомментировать содержимое с упоминанием о том, что первый упомянутый объект имеет атрибут ID 0, а вторая упомянутая сущность имеет атрибут ID 1).|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Получает имя языка и региональных параметров в формате languagecode2-Country/regioncode2 (например, "zh-CN" или "en-US").|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|Определяет формат отображения чисел, соответствующий культуре.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[нумбердеЦималсепаратор](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|Получает строку, используемую в качестве десятичного разделителя для числовых значений.|
||[нумберграупсепаратор](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|Получает строку, используемую для разделения групп цифр слева от десятичного разделителя для числовых значений.|
|[Range](/javascript/api/excel/excel.range)|[moveTo (Дестинатионранже: \| строка Range)](/javascript/api/excel/excel.range#moveto-destinationrange-)|Перемещает значения ячеек, форматирование и формулы из текущего диапазона в конечный диапазон, заменяя старые сведения в этих ячейках.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[Аджустиндент (Amount: число)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Настраивает отступ для форматирования диапазона.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Закрывает текущую книгу.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Сохраняет текущую книгу.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[онровхидденчанжед](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Происходит при изменении скрытого состояния одной или нескольких строк на определенном листе.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|Адрес диапазона, который выполнил вычисление.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[онровхидденчанжед](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Происходит при изменении скрытого состояния одной или нескольких строк на определенном листе.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Получает адрес диапазона, представляющий измененную область конкретного листа.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Получает тип изменения, которое представляет способ запуска события.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Получает идентификатор листа, в котором изменены данные.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
