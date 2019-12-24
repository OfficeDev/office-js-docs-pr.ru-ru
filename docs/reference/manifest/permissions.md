---
title: Элемент Permissions в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: a70d72e454273873c6a30ffd82c3a2a5194f55e0
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851308"
---
# <a name="permissions-element"></a>Элемент Permissions

Указывает уровень доступа к API для надстройки Office. Запрашивая разрешения, руководствуйтесь принципом минимальных привилегий.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

Для надстроек области задач и контентных надстроек:

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

Для почтовых надстроек

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>Содержится в

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Замечания

Более подробную информацию можно узнать [в статье запрос разрешений для использования API в](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) надстройках и [Общие сведения о разрешениях для надстроек Outlook](/outlook/add-ins/understanding-outlook-add-in-permissions).
