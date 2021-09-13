---
title: Объекты window, которые неподтверчены в Office надстройки
description: В этой статье указаны некоторые объекты времени запуска окне, которые не работают в Office надстройки.
ms.date: 07/10/2020
ms.localizationpriority: medium
ms.openlocfilehash: 65cdd4d53dcbcdea75f7eeec39300e4eaee132ac
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154739"
---
# <a name="window-objects-that-are-unsupported-in-office-add-ins"></a>Объекты window, которые неподтверчены в Office надстройки

Для некоторых версий Windows и Office надстройки запускают в internet Explorer 11. (Дополнительные сведения см. в [браузерах, используемых Office надстройки.)](../concepts/browsers-used-by-office-web-add-ins.md) Некоторые свойства или свойства глобального объекта не `window` поддерживаются в Internet Explorer 11. Эти свойства отключены в надстройки, чтобы убедиться, что надстройка обеспечивает согласованный доступ всем пользователям, независимо от того, какой браузер используется надстройка. Это также помогает правильно загружать AngularJS.

Ниже приводится список отключенных свойств. Список находится в процессе выполнения. Если вы обнаружите дополнительные свойства, которые не работают в надстройки, воспользуйтесь ниже средством обратной `window` связи.

- `window.history.pushState`
- `window.history.replaceState`

## <a name="see-also"></a>Дополнительные материалы

- [Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md)