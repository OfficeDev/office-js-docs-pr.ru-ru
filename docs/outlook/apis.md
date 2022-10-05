---
title: API надстроек Outlook
description: Узнайте, как ссылаться на API надстроек Outlook и объявлять разрешения в надстройке Outlook.
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 69043646add5e32502efb0d2a5b1259667e564d9
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467078"
---
# <a name="outlook-add-in-apis"></a>API надстроек Outlook

Чтобы использовать API-интерфейсы в надстройке Outlook, необходимо указать расположение библиотеки Office.js, набор требований, схему и разрешения. В основном вы будете использовать API JavaScript для Office, предоставляемые через объект [Mailbox](#mailbox-object) .

## <a name="officejs-library"></a>Библиотека Office.js

Для взаимодействия с [API надстройки Outlook необходимо использовать API](/javascript/api/outlook) JavaScript в Office.js. Сеть доставки содержимого (CDN) для библиотеки имеет значение `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`. Надстройки, отправляемые в AppSource, должны ссылаться на библиотеку Office.js в этой сети CDN. Нельзя использовать локальную ссылку.

Добавьте ссылку на сеть CDN в теге `<script>`, вложенном в тег `<head>`, на веб-странице (HTML-, ASPX- или PHP-файле), где определяется пользовательский интерфейс надстройки.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

As we add new APIs, the URL to Office.js will stay the same. We will change the version in the URL only if we break an existing API behavior.

> [!IMPORTANT]
> При разработке надстройки для любого клиентского приложения Office используйте ссылку на API `<head>` JavaScript для Office из раздела страницы. Это гарантирует, что API полностью инициализируется раньше всех элементов body.

## <a name="requirement-sets"></a>Наборы обязательных элементов

Все API Outlook относятся к [набору обязательных элементов почтового ящика](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets). Существуют разные версии набора обязательных элементов `Mailbox`. Каждый новый набор API, который мы выпускаем, относится к более высокой версии набора. Не все клиенты Outlook поддерживают самый новый набор API при их выпуске, но если для клиента Outlook объявлена поддержка определенного набора обязательных элементов, то он поддерживает все API из этого набора.

To control which Outlook clients the add-in appears in, specify a minimum requirement set version in the manifest. For example, if you specify requirement set version 1.3, the add-in will not show up in any Outlook client that doesn't support a minimum version of 1.3.

Specifying a requirement set doesn't limit your add-in to the APIs in that version. If the add-in specifies requirement set v1.1 but is running in an Outlook client that supports v1.3, the add-in can still use v1.3 APIs. The requirement set only controls which Outlook clients the add-in appears in.

Чтобы проверить доступность API-интерфейсов из набора требований, версия которого выше указанной в манифесте, вы можете использовать стандартный код JavaScript:

```js
if (item.somePropertyOrFunction) {
   item.somePropertyOrFunction...  
}
```

> [!NOTE]
> Эти проверки необязательны для API, входящих в набор обязательных элементов версии, указанной в манифесте.

Укажите минимальный набор обязательных элементов API для вашего сценария, без которого функции надстройки не будут работать. Набор обязательных элементов указывается в манифесте. Разметка зависит от манифеста, который вы используете. 

- **XML-манифест**: используйте **\<Requirements\>** элемент. Обратите внимание, **\<Methods\>** что дочерний элемент **\<Requirements\>** надстроек Outlook не поддерживается, поэтому невозможно объявить поддержку определенных методов.
- **Манифест Teams (предварительная версия)**: используйте свойство extensions.capabilities. 

Дополнительные сведения см. в манифестах надстроек [Outlook](manifests.md) и наборах обязательных элементов [API Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets).

## <a name="permissions"></a>Разрешения

Для использования необходимых API-интерфейсов надстройке нужны соответствующие разрешения. Как правило, следует указывать минимальные разрешения, необходимые надстройке.

Существует четыре уровня разрешений. **ограниченный**, **чтение,** **чтение и запись**, а **также почтовый ящик для чтения и записи**. Дополнительные сведения. Дополнительные сведения см. в [разделе "Общие сведения о разрешениях надстроек Outlook"](understanding-outlook-add-in-permissions.md).

## <a name="mailbox-object"></a>Объект Mailbox

[!include[information about Mailbox object](../includes/mailbox-object-desc.md)]

## <a name="see-also"></a>См. также

- [Манифесты надстроек Outlook](manifests.md)
- [Общие сведения о наборах требований API Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)
- [Общие сведения о разрешениях надстроек Outlook](understanding-outlook-add-in-permissions.md).
- [Конфиденциальность и безопасность надстроек Office](../concepts/privacy-and-security.md)
