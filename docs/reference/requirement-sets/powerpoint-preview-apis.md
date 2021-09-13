---
title: PowerPoint API предварительного просмотра JavaScript
description: Сведения о предстоящих PowerPoint API JavaScript.
ms.date: 01/27/2021
ms.prod: powerpoint
ms.localizationpriority: medium
ms.openlocfilehash: d9cb28c56a84829d87ba30e494aa46b927e0bc64
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154896"
---
# <a name="powerpoint-javascript-preview-apis"></a>PowerPoint API предварительного просмотра JavaScript

Новые PowerPoint API JavaScript сначала вводятся в "предварительную версию", а затем становятся частью определенного набора требований с номерами после достаточного тестирования и получения отзывов пользователей.

В первой таблице представлен краткий обзор API, а в последующей таблице приведен подробный список.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| Управление слайдами | Добавляет поддержку для добавления слайдов, а также управления макетами слайдов и мастерами слайдов. | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| Фигуры | Добавляет поддержку для получения ссылок на фигуры на слайде. | [Shape](/javascript/api/powerpoint/powerpoint.shape) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены PowerPoint API JavaScript, которые в настоящее время находятся в предварительном просмотре. Полный список всех API PowerPoint JavaScript (включая API предварительного просмотра и ранее выпущенные API), см. Excel [API JavaScript.](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutId)|Указывает ID макета слайда, который будет использоваться для нового слайда.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slideMasterId)|Указывает ID мастера слайдов, который будет использоваться для нового слайда.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slideMasters)|Возвращает коллекцию `SlideMaster` объектов, которые находятся в презентации.|
||[tags](/javascript/api/powerpoint/powerpoint.presentation#tags)|Возвращает коллекцию тегов, присоединенных к презентации.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[delete()](/javascript/api/powerpoint/powerpoint.shape#delete__)|Удаляет фигуру из коллекции фигур.|
||[id](/javascript/api/powerpoint/powerpoint.shape#id)|Получает уникальный ID формы.|
||[tags](/javascript/api/powerpoint/powerpoint.shape#tags)|Возвращает коллекцию тегов в форме.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getCount__)|Получает количество фигур в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getItem_key_)|Получает форму с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemAt_index_)|Получает фигуру с помощью нулевого индекса в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemOrNullObject_id_)|Получает форму с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[макет](/javascript/api/powerpoint/powerpoint.slide#layout)|Получает макет слайда.|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#shapes)|Возвращает коллекцию фигур на слайде.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slideMaster)|Получает `SlideMaster` объект, который представляет содержимое слайда по умолчанию.|
||[tags](/javascript/api/powerpoint/powerpoint.slide#tags)|Возвращает коллекцию тегов на слайде.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint. AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_)|Добавляет новый слайд в конце коллекции.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|Получает уникальный ID макета слайда.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|Получает имя макета слайда.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getCount__)|Получает количество макетов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItem_key_)|Получает макет с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemAt_index_)|Получает макет с использованием нулевого индекса в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemOrNullObject_id_)|Получает макет с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|Получает уникальный ID мастера слайдов.|
||[макеты](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|Получает коллекцию макетов, предоставленных мастером слайдов для слайдов.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|Получает уникальное имя мастера слайдов.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getCount__)|Получает число мастеров слайдов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItem_key_)|Получает мастер слайдов с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemAt_index_)|Получает мастер слайдов с помощью нулевого индекса в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemOrNullObject_id_)|Получает мастер слайдов с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#key)|Получает уникальный ID тега.|
||[value](/javascript/api/powerpoint/powerpoint.tag#value)|Получает значение тега.|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_)|Добавляет новый тег в конце коллекции.|
||[delete (key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#delete_key_)|Удаляет тег с заданным `key` в этой коллекции.|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#getCount__)|Получает количество тегов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItem_key_)|Получает тег с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemAt_index_)|Получает тег с использованием нулевого индекса в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemOrNullObject_key_)|Получает тег с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>Дополнительные материалы

- [PowerPoint Справочная документация по API JavaScript](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для PowerPoint](powerpoint-api-requirement-sets.md)