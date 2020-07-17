---
title: Объекты Window, которые не поддерживаются в надстройках Office
description: В этой статье указаны некоторые объекты среды выполнения, которые не работают в надстройках Office.
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: d2560748841bd1e2a7708b25a8e51133563d1534
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160507"
---
# <a name="window-objects-that-are-unsupported-in-office-add-ins"></a>Объекты Window, которые не поддерживаются в надстройках Office

В некоторых версиях Windows и Office надстройки запускаются в среде выполнения Internet Explorer 11. (Дополнительные сведения см. в разделе [браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md).) Некоторые свойства или вложенные свойства глобального `window` объекта не поддерживаются в Internet Explorer 11. Эти свойства отключены в надстройках, чтобы надстройка гарантированно соответствовала всем пользователям, независимо от того, какой браузер использует надстройка. Это также способствует правильной загрузке AngularJS.

Ниже приведен список отключенных свойств. Список является выполняемой работой. Если вы обнаружите дополнительные `window` свойства, которые не работают в надстройках, воспользуйтесь средством обратной связи, приведенным ниже, чтобы сообщить нам об этом.

- `window.history.pushState`
- `window.history.replaceState`

## <a name="see-also"></a>См. также

- [Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md)