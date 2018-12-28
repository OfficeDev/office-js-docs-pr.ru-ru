---
title: Добавление ссылок на библиотеку JavaScript API для Office из сети доставки содержимого (CDN)
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 9943af86419652e5f5e89b1741b32b4e0da15e77
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457623"
---
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a>Добавление ссылок на библиотеку JavaScript API для Office из сети доставки содержимого (CDN)

> [!NOTE]
> Помимо действий, описанных в этой статье, если необходимо использовать TypeScript, для получения Intellisense потребуется запустить указанную ниже команду в системном приглашении с поддержкой Node (или в окне Git Bash) из корневой папки проекта. У вас должен быть установлен [Node.js](https://nodejs.org) (содержащий npm).
> 
> ```bash
> npm install --save-dev @types/office-js
> ```

Библиотека [API JavaScript для Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) состоит из файла Office.js и связанных JS-файлов ведущего приложения, например Excel-15.js и Outlook-15.js. 


Простейший способ добавить ссылку на API — использовать нашу сеть доставки содержимого (CDN), добавив следующий код `<script>` в тег `<head>` страницы:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

`/1/` перед `office.js` в URL-адресе CDN указывает, что необходимо использовать последний добавочный выпуск файла Office.js версии 1. Так как API JavaScript для Office обеспечивает обратную совместимость, в последнем выпуске будут и дальше поддерживаться элементы API, представленные ранее в версии 1. Если вам нужно обновить существующий проект, см. статью [Обновление версии API JavaScript для Office и файлов схемы манифеста](update-your-javascript-api-for-office-and-manifest-schema-version.md). 

Если вы планируете опубликовать свою надстройку Office из AppSource, необходимо использовать эту ссылку на сеть CDN. Локальные ссылки подходят только для внутренних сценариев, а также сценариев разработки и отладки.

> [!IMPORTANT]
>  Разрабатывая надстройку для ведущего приложения Office, ссылайтесь на API JavaScript для Office из раздела `<head>` страницы. Это гарантирует, что API полностью инициализируется раньше всех элементов основного текста. Ведущим приложениям Office необходимо, чтобы надстройки инициализировались в течение 5 секунд после активации. Если надстройка не активируется в течение этого срока, будет считаться, что она не отвечает, и пользователь увидит сообщение об ошибке.       

## <a name="see-also"></a>См. также

- [Общие сведения об API JavaScript для Office](understanding-the-javascript-api-for-office.md)    
- [API JavaScript для Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
    
