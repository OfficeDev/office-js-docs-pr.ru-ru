> [!NOTE]
> API предварительной версии могут быть изменены и не предназначены для использования в рабочей среде. Рекомендуется использовать их только в тестовой среде и среде разработки. Не используйте API предварительной версии в рабочей среде или в важных деловых документах.
>
> Для использования API предварительной версии:
>
> - Необходимо ссылаться на **бета-** библиотеку в сети CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) . [Файл определения типа](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) для компиляции TypeScript и IntelliSense находится в сети CDN и [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Вы можете установить эти типы с помощью `npm install --save-dev @types/office-js-preview` .
> - Вам может потребоваться присоединиться к [программе предварительной оценки Office](https://insider.office.com) для получения доступа к более последним сборкам Office.