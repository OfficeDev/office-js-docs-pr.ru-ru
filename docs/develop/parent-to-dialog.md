---
title: Альтернативные способы передачи сообщений в диалоговое окно со своей хост-страницы
description: Узнайте обходные пути, которые можно использовать, если метод messageChild не поддерживается.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: d664178a804b206e02634326cc27699fc6ceb0f7
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58939326"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a>Альтернативные способы передачи сообщений в диалоговое окно со своей хост-страницы

Рекомендуемый способ передачи данных и сообщений с родительской страницы в диалоговое окно для детей используется метод, описанный в API диалоговых Office в Office надстройки `messageChild` . [](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) Если надстройка работает на платформе или хосте, не поддерживающей набор требований [DialogApi 1.2,](../reference/requirement-sets/dialog-api-requirement-sets.md)существует два других способа передать информацию в диалоговое окно.

- Добавьте параметры запроса в URL-адрес, который передается в метод `displayDialogAsync`.
- Храните информацию в месте, доступном как для главного, так и для диалогового окна. Два окна не имеют общего хранилища сеансов (свойство [Window.sessionStorage),](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) но если они имеют один и тот же домен *(включая* номер [порта,](https://www.w3schools.com/html/html5_webstorage.asp)если таковые имеются), они имеют общий локальный служба хранилища .\*

> [!NOTE]
> \* Существует ошибка, влияющая на вашу стратегию обработки маркеров. Если надстройка работает в **Office в Интернете** с использованием браузера Safari или Microsoft Edge, у диалогового окна и области задач нет одного общего локального хранилища, поэтому его нельзя использовать для связи между ними.

## <a name="use-local-storage"></a>Использование локального хранилища

Чтобы использовать локальное хранилище, перед вызовом вызываем метод объекта на хост-странице, `setItem` `window.localStorage` как в следующем `displayDialogAsync` примере.

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

Код в диалоговом окне читает элемент при необходимости, как в следующем примере.

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

## <a name="use-query-parameters"></a>Использование параметров запроса

В приведенном ниже примере показано, как передавать данные с помощью параметра запроса.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

Пример, в котором используется эта техника, см. в статье [Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

Код в вашем диалоговом окне может проанализировать URL-адрес и прочитать значение параметра.

> [!IMPORTANT]
> Office автоматически добавляет параметр запроса `_host_info` в URL-адрес, который передается `displayDialogAsync`. (Этот параметр добавляется после пользовательских параметров запроса, если они есть. Он не добавляется в последующие URL-адреса, которые открываются в диалоговом окне.) Корпорация Майкрософт может изменить содержимое этого значения или удалить его полностью, поэтому ваш код не должен его считывать. Это же значение добавляется в хранилище сеансов диалоговое окно (свойство [Window.sessionStorage).](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) *Ваш код не должен ни считывать это значение, ни записывать в него данные*.
