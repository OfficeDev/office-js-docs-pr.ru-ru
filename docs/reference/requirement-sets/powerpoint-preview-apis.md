---
title: API предварительного просмотра PowerPoint JavaScript
description: Сведения о предстоящих API JavaScript PowerPoint.
ms.date: 01/27/2021
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: 042ce0c2b42b2c0dca9900982376cd568a4a3622
ms.sourcegitcommit: 929dcf2f415b94f42330a9035ed11a5cedad88f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/16/2021
ms.locfileid: "50830974"
---
# <a name="powerpoint-javascript-preview-apis"></a>API предварительного просмотра PowerPoint JavaScript

Новые API JavaScript PowerPoint сначала вводятся в "предварительную версию", а затем становятся частью определенного набора требований с номерами после достаточного тестирования и получения отзывов пользователей.

В первой таблице представлен краткий обзор API, а в последующей таблице приведен подробный список.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| Управление слайдами | Добавляет поддержку для добавления слайдов, а также управления макетами слайдов и мастерами слайдов. | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| Фигуры | Добавляет поддержку для получения ссылок на фигуры на слайде. | [Shape](/javascript/api/powerpoint/powerpoint.shape) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API JavaScript PowerPoint, которые в настоящее время находятся в предварительном просмотре. Полный список всех API JavaScript PowerPoint (включая API предварительного просмотра и ранее выпущенные API) см. во всех API [JavaScript Excel.](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutid)|Указывает ID макета слайда, который будет использоваться для нового слайда.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slidemasterid)|Указывает ID мастера слайдов, который будет использоваться для нового слайда.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slidemasters)|Возвращает коллекцию `SlideMaster` объектов, которые находятся в презентации.|
||[tags](/javascript/api/powerpoint/powerpoint.presentation#tags)|Возвращает коллекцию тегов, присоединенных к презентации.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[id](/javascript/api/powerpoint/powerpoint.shape#id)|Получает уникальный ID формы.|
||[tags](/javascript/api/powerpoint/powerpoint.shape#tags)|Возвращает коллекцию тегов в форме.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getcount--)|Получает количество фигур в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitem-key-)|Получает форму с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemat-index-)|Получает фигуру с помощью нулевого индекса в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemornullobject-id-)|Получает форму с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[макет](/javascript/api/powerpoint/powerpoint.slide#layout)|Получает макет слайда.|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#shapes)|Возвращает коллекцию фигур на слайде.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slidemaster)|Получает `SlideMaster` объект, который представляет содержимое слайда по умолчанию.|
||[tags](/javascript/api/powerpoint/powerpoint.slide#tags)|Возвращает коллекцию тегов на слайде.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint.AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add-options-)|Добавляет новый слайд в конце коллекции.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|Получает уникальный ID макета слайда.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|Получает имя макета слайда.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getcount--)|Получает количество макетов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitem-key-)|Получает макет с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemat-index-)|Получает макет с использованием нулевого индекса в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemornullobject-id-)|Получает макет с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|Получает уникальный ID мастера слайдов.|
||[макеты](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|Получает коллекцию макетов, предоставленных мастером слайдов для слайдов.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|Получает уникальное имя мастера слайдов.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getcount--)|Получает число мастеров слайдов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitem-key-)|Получает мастер слайдов с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemat-index-)|Получает мастер слайдов с помощью нулевого индекса в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemornullobject-id-)|Получает мастер слайдов с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#key)|Получает уникальный ID тега.|
||[value](/javascript/api/powerpoint/powerpoint.tag#value)|Получает значение тега.|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#add-key--value-)|Добавляет новый тег в конце коллекции.|
||[delete (key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#delete-key-)|Удаляет тег с заданным `key` в этой коллекции.|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#getcount--)|Получает количество тегов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getitem-key-)|Получает тег с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#getitemat-index-)|Получает тег с использованием нулевого индекса в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getitemornullobject-key-)|Получает тег с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript PowerPoint](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для PowerPoint](powerpoint-api-requirement-sets.md)