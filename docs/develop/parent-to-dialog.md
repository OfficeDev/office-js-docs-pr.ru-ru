---
title: Альтернативные способы передачи сообщений в диалоговое окно со своей хост-страницы
description: Узнайте обходные пути, которые необходимо использовать, если метод messageChild не поддерживается.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 94b002970084a13fe69e1469aa26577039600989
ms.sourcegitcommit: 287a58de82a09deeef794c2aa4f32280efbbe54a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/28/2022
ms.locfileid: "64496812"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a>Альтернативные способы передачи сообщений в диалоговое окно со своей хост-страницы

Рекомендуемый `messageChild` способ передачи данных и сообщений с родительской страницы в диалоговое окно с помощью метода, описанного в [API](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) диалоговых Office в Office надстройки. Если надстройка работает на платформе или хосте, не поддерживающей набор требований [DialogApi 1.2](/javascript/api/requirement-sets/common/dialog-api-requirement-sets), существует два других способа передать информацию в диалоговое окно.

- Добавьте параметры запроса в URL-адрес, который передается в метод `displayDialogAsync`.
- Храните информацию в месте, доступном как для главного, так и для диалогового окна. Два окна не имеют общего хранилища сеансов (свойство [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)), но если они имеют один и тот же *домен (включая* номер порта, если таковые имеются), они имеют общий локальный [служба хранилища](https://www.w3schools.com/html/html5_webstorage.asp).\*

> [!NOTE]
> \* Существует ошибка, влияющая на вашу стратегию обработки маркеров. Если надстройка работает в **Office в Интернете** с использованием браузера Safari или Microsoft Edge, у диалогового окна и области задач нет одного общего локального хранилища, поэтому его нельзя использовать для связи между ними.

## <a name="use-local-storage"></a>Использование локального хранилища

Чтобы использовать локальное хранилище, `setItem` `window.localStorage` `displayDialogAsync` перед вызовом вызываем метод объекта на хост-странице, как в следующем примере.

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
> Office автоматически добавляет параметр запроса `_host_info` в URL-адрес, который передается `displayDialogAsync`. (Этот параметр добавляется после пользовательских параметров запроса, если они есть. Он не добавляется в последующие URL-адреса, которые открываются в диалоговом окне.) Корпорация Майкрософт может изменить содержимое этого значения или удалить его полностью, поэтому ваш код не должен его считывать. Это же значение добавляется в хранилище сеансов диалоговое окно (свойство [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ). *Ваш код не должен ни считывать это значение, ни записывать в него данные*.
