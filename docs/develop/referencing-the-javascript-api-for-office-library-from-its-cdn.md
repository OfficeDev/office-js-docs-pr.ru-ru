---
title: 'Ссылки на библиотеку API JavaScript для Office '
description: Узнайте, как ссылаться на Office API JavaScript и определения типов в надстройки.
ms.date: 02/18/2021
localization_priority: Normal
ms.openlocfilehash: 04f97412c07cb39f5b2f753c3ce14e56e87c3de5
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938636"
---
# <a name="referencing-the-office-javascript-api-library"></a>Ссылки на библиотеку API JavaScript для Office 

Библиотека [API Office JavaScript предоставляет API,](../reference/javascript-api-for-office.md) которые ваша надстройка может использовать для взаимодействия с Office приложением. Самый простой способ ссылки на библиотеку — использовать сеть доставки контента (CDN) путем добавления следующего тега в разделе `<script>` `<head>` вашей HTML-страницы.

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

При первом загрузке надстройки и кэшировании Office API JavaScript, чтобы убедиться, что она использует самые последние реализации Office.js и связанных с ними файлов для указанной версии.

> [!IMPORTANT]
> Необходимо ссылаться на Office API JavaScript из раздела страницы, чтобы убедиться, что API полностью инициализирован перед `<head>` любыми элементами тела.

## <a name="api-versioning-and-backward-compatibility"></a>Версия API и обратная совместимость

В предыдущем фрагменте HTML перед URL-адресом CDN указывается последний дополнительный выпуск в версии `/1/` `office.js` 1 Office.js. Поскольку API Office JavaScript поддерживает обратную совместимость, последний выпуск будет по-прежнему поддерживать членов API, которые были представлены ранее в версии 1. Если вам необходимо обновить существующий проект, см. в статью Обновление версии API javaScript Office JavaScript и [файлы схемы манифеста.](update-your-javascript-api-for-office-and-manifest-schema-version.md) 

Если вы планируете опубликовать свою надстройку Office из AppSource, необходимо использовать эту ссылку на сеть CDN. Локальные ссылки подходят только для внутренних сценариев, а также сценариев разработки и отладки.

> [!NOTE]
> Чтобы использовать API предварительных версий, требуется указать ссылку на предварительную версию библиотеки API JavaScript для Office в сети CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.

## <a name="enabling-intellisense-for-a-typescript-project"></a>Включение IntelliSense для проекта TypeScript

В дополнение к ссылке на API Office JavaScript, как описано выше, вы также можете включить IntelliSense для проекта надстройки TypeScript с помощью определений типа из [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js). Для этого запустите следующую команду в системном запросе с поддержкой узла (или в окне баш git) из корневой папки проекта. У вас должен быть установлен [Node.js](https://nodejs.org) (содержащий npm).

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a>Предварительные API

Новые API JavaScript сначала вводятся в "предварительную версию", а затем становятся частью определенного набора требований с номерами после достаточного тестирования и необходимости отзыва пользователей.

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a>См. также

- [Общие сведения об API JavaScript для Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript для Office](../reference/javascript-api-for-office.md)
