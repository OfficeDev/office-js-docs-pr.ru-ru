---
title: Альтернативные способы передачи сообщений в диалоговое окно с главной страницы
description: Узнайте, как использовать методы обхода, если метод Мессажечилд не поддерживается.
ms.date: 09/24/2020
localization_priority: Normal
ms.openlocfilehash: 8f44f7f5c145b58d13e7387d01e28fd349a512fc
ms.sourcegitcommit: b47318a24a50443b0579e05e178b3bb5433c372f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/25/2020
ms.locfileid: "48279486"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a>Альтернативные способы передачи сообщений в диалоговое окно с главной страницы

Рекомендуемый способ передачи данных и сообщений из родительской страницы в дочернее диалоговое окно осуществляется с помощью `messageChild` метода, как описано в статье [Использование API диалоговых окон Office в](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)надстройках Office. Если ваша надстройка работает на платформе или узле, которая не поддерживает [набор требований DialogApi 1,2](../reference/requirement-sets/dialog-api-requirement-sets.md), существует два других способа передачи сведений в диалоговое окно:

- Добавьте параметры запроса в URL-адрес, который передается в метод `displayDialogAsync`.
- Храните информацию в месте, доступном как для главного, так и для диалогового окна. Два окна не используют общее хранилище сеанса (свойство [Window. sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ), но *если они имеют один и тот же домен* (включая номер порта, если они есть), они совместно используют общее [Локальное хранилище](https://www.w3schools.com/html/html5_webstorage.asp).\*


> [!NOTE]
> \* Существует ошибка, влияющая на вашу стратегию обработки маркеров. Если надстройка работает в **Office в Интернете** с использованием браузера Safari или Microsoft Edge, у диалогового окна и области задач нет одного общего локального хранилища, поэтому его нельзя использовать для связи между ними.

## <a name="use-local-storage"></a>Использование локального хранилища

Чтобы использовать локальное хранилище, вызовите `setItem` метод `window.localStorage` объекта на главной странице перед `displayDialogAsync` вызовом, как показано в следующем примере:

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

Код в диалоговом окне считывает элемент, когда он необходим, как в следующем примере:

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
> Office автоматически добавляет параметр запроса `_host_info` в URL-адрес, который передается `displayDialogAsync`. (Этот параметр добавляется после пользовательских параметров запроса, если они есть. Он не добавляется в последующие URL-адреса, которые открываются в диалоговом окне.) Корпорация Майкрософт может изменить содержимое этого значения или удалить его полностью, поэтому ваш код не должен его считывать. Одно и то же значение добавляется в хранилище сеанса диалогового окна (свойство [Window. sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ). *Ваш код не должен ни считывать это значение, ни записывать в него данные*.
