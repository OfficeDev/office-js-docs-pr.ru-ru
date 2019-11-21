---
title: Набор обязательных элементов API JavaScript для Excel Online
description: Сведения о наборе требований Ексцелапионлине
ms.date: 11/19/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: e583c9832f04e17dc1c82d38d056fe2749888a77
ms.sourcegitcommit: e56bd8f1260c73daf33272a30dc5af242452594f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/21/2019
ms.locfileid: "38757494"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Набор обязательных элементов API JavaScript для Excel Online

Набор `ExcelApiOnline` требований является особым набором требований, включающим функции, доступные только для Excel в Интернете. API в этом наборе обязательных элементов считаются рабочими API (не подчиняются недокументированным изменениям поведения или структурным изменениям) для Excel на веб-узле. `ExcelApiOnline`считаются API-интерфейсами Preview для других платформ (Windows, Mac, iOS) и могут не поддерживаться ни одной из этих платформ.

Если API в `ExcelApiOnline` наборе обязательных элементов поддерживаются на всех платформах, они будут добавлены к следующему набору`ExcelApi 1.[NEXT]`обязательных требований (). После того как новое требование будет общедоступным, эти API будут `ExcelApiOnline`удалены из. Это можно считать похожим процессом повышения роли API, который перемещается из бета-версии в выпуск.

> [!IMPORTANT]
> `ExcelApiOnline`является надмножеством набора последних пронумерованных требований.

> [!IMPORTANT]
> `ExcelApiOnline 1.1`— Единственная версия интерфейсов API, предназначенных только для интерактивного подключения. Это связано с тем, что Excel в Интернете всегда будет иметь одну версию, доступную для пользователей с последней версией.

## <a name="recommended-usage"></a>Рекомендуемое использование

Так `ExcelApiOnline` как API поддерживаются только в Excel в Интернете, надстройка должна проверить, поддерживается ли набор требований, прежде чем вызывать эти API. Это позволяет избежать вызова Интернет-API на другой платформе.

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

После того как API находится в наборе требований к нескольким платформам, необходимо удалить или изменить `isSetSupported` проверку. При этом функция надстройки будет включена на других платформах. При внесении изменений обязательно протестируйте эту функцию на этих платформах.

> [!IMPORTANT]
> В манифесте невозможно `ExcelApiOnline 1.1` указать требования для активации. Значение не является допустимым для использования в [элементе Set](../manifest/set.md).

## <a name="api-list"></a>Список API

В настоящее время интерфейсы API, доступные только в Интернете, отсутствуют. Возврат к последующим добавлению новых компонентов в Excel в Интернете и поддержка API JavaScript для Office.

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-online)
- [Предварительные версии API JavaScript для Excel](./excel-preview-apis.md)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)