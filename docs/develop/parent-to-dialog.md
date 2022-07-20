---
title: Альтернативные способы передачи сообщений в диалоговое окно со страницы узла
description: Узнайте, как использовать обходные пути, если метод messageChild не поддерживается.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: f42a549a3c39866516cfd5395dd7589a890b0956
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889417"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a>Альтернативные способы передачи сообщений в диалоговое окно со страницы узла

Рекомендуемый `messageChild` способ передачи данных и сообщений с родительской страницы в дочернее диалоговое окно — это метод, описанный в разделе "Использование [API](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) диалогового окна Office" в надстройке Office. Если надстройка работает на платформе или узле, не поддерживающих набор обязательных элементов [DialogApi 1.2](/javascript/api/requirement-sets/common/dialog-api-requirement-sets), существует два других способа передачи информации в диалоговое окно.

- Добавьте параметры запроса в URL-адрес, который передается в метод `displayDialogAsync`.
- Храните информацию в месте, доступном как для главного, так и для диалогового окна. Два окна не совместно используют общее хранилище сеансов (свойство [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ), но если у них один и тот же *домен (включая* номер порта, если он есть), они совместно используют общее локальное [хранилище](https://www.w3schools.com/html/html5_webstorage.asp).\*

> [!NOTE]
> \* Существует ошибка, влияющая на вашу стратегию обработки маркеров. Если надстройка работает в **Office в Интернете** с использованием браузера Safari или Microsoft Edge, у диалогового окна и области задач нет одного общего локального хранилища, поэтому его нельзя использовать для связи между ними.

## <a name="use-local-storage"></a>Использование локального хранилища

Чтобы использовать локальное хранилище, вызовите `setItem` `window.localStorage` `displayDialogAsync` метод объекта на странице узла перед вызовом, как показано в следующем примере.

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

Код в диалоговом окне считывает элемент при необходимости, как показано в следующем примере.

```js
const clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// const clientID = localStorage.clientID;
```

## <a name="use-query-parameters"></a>Использование параметров запроса

В приведенном ниже примере показано, как передавать данные с помощью параметра запроса.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

Пример, в котором используется эта техника, см. в статье [Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

Код в вашем диалоговом окне может проанализировать URL-адрес и прочитать значение параметра.

> [!IMPORTANT]
> Office автоматически добавляет параметр запроса `_host_info` в URL-адрес, который передается `displayDialogAsync`. (Этот параметр добавляется после пользовательских параметров запроса, если они есть. Он не добавляется в последующие URL-адреса, которые открываются в диалоговом окне.) Корпорация Майкрософт может изменить содержимое этого значения или удалить его полностью, поэтому ваш код не должен его считывать. Это же значение добавляется в хранилище сеансов диалогового окна (свойство [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ). *Ваш код не должен ни считывать это значение, ни записывать в него данные*.
