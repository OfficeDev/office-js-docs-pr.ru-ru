---
title: 'Ссылки на библиотеку API JavaScript для Office '
description: Узнайте, как ссылаться на библиотеку API JavaScript Office и определения типов в надстройки.
ms.date: 02/18/2021
localization_priority: Normal
ms.openlocfilehash: 346a34c0cbc31b5e569a5106dcd2bc01593b114a
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505194"
---
# <a name="referencing-the-office-javascript-api-library"></a>Ссылки на библиотеку API JavaScript для Office 

Библиотека [API JavaScript](../reference/javascript-api-for-office.md) Office предоставляет API, которые ваша надстройка может использовать для взаимодействия с приложением Office. Самый простой способ ссылки на библиотеку — использовать сеть доставки контента (CDN), добавив следующий тег в разделе `<script>` `<head>` вашей HTML-страницы:  

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

При первом загрузке надстройки будут загружаться и кэшируются файлы API Office JavaScript, чтобы убедиться, что она использует самые последние реализации Office.js и связанных с ними файлов для указанной версии.

> [!IMPORTANT]
> Чтобы убедиться, что API полностью инициализирован перед любыми элементами тела, необходимо ссылаться на API JavaScript Office из раздела `<head>` страницы.

## <a name="api-versioning-and-backward-compatibility"></a>Версия API и обратная совместимость

В предыдущем фрагменте HTML перед URL-адресом CDN указывается последний дополнительный выпуск в версии `/1/` `office.js` 1 Office.js. Так как API JavaScript Office поддерживает обратную совместимость, последний выпуск будет по-прежнему поддерживать участников API, которые были представлены ранее в версии 1. Если необходимо обновить существующий проект, см. в статью Обновление версии [API JavaScript Office и файлы схемы манифеста.](update-your-javascript-api-for-office-and-manifest-schema-version.md) 

Если вы планируете опубликовать свою надстройку Office из AppSource, необходимо использовать эту ссылку на сеть CDN. Локальные ссылки подходят только для внутренних сценариев, а также сценариев разработки и отладки.

> [!NOTE]
> Чтобы использовать API предварительных версий, требуется указать ссылку на предварительную версию библиотеки API JavaScript для Office в сети CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.

## <a name="enabling-intellisense-for-a-typescript-project"></a>Включение IntelliSense для проекта TypeScript

Помимо ссылок на API JavaScript Office, как описано выше, вы также можете включить IntelliSense для проекта надстройки TypeScript с помощью определений типа [из DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js). Для этого запустите следующую команду в системном запросе с поддержкой узла (или в окне баш git) из корневой папки проекта. У вас должен быть установлен [Node.js](https://nodejs.org) (содержащий npm).

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a>Предварительные API

Новые API JavaScript сначала вводятся в "предварительную версию", а затем становятся частью определенного набора требований с номерами после достаточного тестирования и необходимости отзыва пользователей.

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a>См. также

- [Общие сведения об API JavaScript для Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript для Office](../reference/javascript-api-for-office.md)
