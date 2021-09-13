---
title: Элемент Permissions в файле манифеста
description: Элемент Permissions указывает уровень доступа к API для Office надстройки.
ms.date: 06/26/2020
ms.localizationpriority: medium
ms.openlocfilehash: a472d7a6f375c3a171fdd529b993aaf2c6109ce9
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154853"
---
# <a name="permissions-element"></a>Элемент Permissions

Указывает уровень доступа к API для надстройки Office. Запрашивая разрешения, руководствуйтесь принципом минимальных привилегий.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

Для надстроек области задач и контентных надстроек:

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

Для почтовых надстроек:

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>Содержится в

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Замечания

Дополнительные сведения см. в материале Запрашивание разрешений на использование [API](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) в надстройке контента и области задач и Outlook разрешений [надстройки.](../../outlook/understanding-outlook-add-in-permissions.md)
