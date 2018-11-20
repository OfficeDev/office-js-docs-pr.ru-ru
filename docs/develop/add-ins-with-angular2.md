---
title: Разработка надстроек Office с помощью Angular
description: ''
ms.date: 11/02/2018
ms.openlocfilehash: 312317e594024125e2dc86d23840750e48d81e40
ms.sourcegitcommit: c6723a31b48945ca4c466ba016a3dfc7b6267f5c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/03/2018
ms.locfileid: "25942253"
---
# <a name="develop-office-add-ins-with-angular"></a>Разработка надстроек Office с помощью Angular

В этой статье приведены рекомендации по использованию Angular 2 и более поздних версий для создания надстройки Office в виде одностраничного приложения.

> [!NOTE]
> Вы можете поделиться опытом по использованию Angular для создания надстроек Office? Примите участие в создании этой статьи на сайте [GitHub](https://github.com/OfficeDev/office-js-docs) или сообщите о [проблеме](https://github.com/OfficeDev/office-js-docs-pr/issues) в соответствующем репозитории. 

Пример надстройки Office, созданной на платформе Angular, приведен в статье [Надстройка на основе Angular для проверки стиля в приложении Word](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).

## <a name="install-the-typescript-type-definitions"></a>Установка определений типов TypeScript
Откройте окно nodejs и введите в командной строке следующую команду: 

```bash
npm install --save-dev @types/office-js
```

## <a name="bootstrapping-must-be-inside-officeinitialize"></a>Начальная загрузка должна определяться в методе Office.initialize

На любой странице, которая вызывает интерфейсы API JavaScript для Office, Word или Excel, в коде сначала нужно назначить метод для свойства `Office.initialize`. (Если у вас нет кода инициализации, тело метода может состоять из пустых символов "`{}`", но свойство `Office.initialize` должно быть определено. Дополнительные сведения см. в разделе [Инициализация надстройки](understanding-the-javascript-api-for-office.md#initializing-your-add-in).) Office вызывает этот метод сразу же после того, как инициализирует библиотеки JavaScript для Office.

**Вызов кода начальной загрузки на основе Angular необходимо задать в методе, который назначен `Office.initialize`**, чтобы сначала выполнялась инициализация библиотек JavaScript для Office. Вот простой пример, в котором показано, как это сделать. Этот код должен находиться в файле main.ts проекта.

```js
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { AppModule } from './app.module';

Office.initialize = function () {
  const platform = platformBrowserDynamic();
  platform.bootstrapModule(AppModule);
};
```

## <a name="use-the-hash-location-strategy-in-the-angular-application"></a>Используйте стратегию навигации с помощью хэша в приложении на основе Angular

Переход между маршрутами в приложении может не выполняться, если не задать стратегию навигации с помощью хэша. Это можно сделать одним из двух способов. Способ первый: указать поставщика стратегии навигации в модуле приложения, как показано в приведенном ниже примере. (Это для файла app.module.ts.)

```js
import { LocationStrategy, HashLocationStrategy } from '@angular/common';
// Other imports suppressed for brevity

@NgModule({
  providers: [
    { provide: LocationStrategy, useClass: HashLocationStrategy },
    // Other providers suppressed
  ],
  // Other module properties suppressed
})
export class AppModule { }
``` 

Если маршруты определены в отдельном модуле маршрутизации, можно задать стратегию навигации с помощью хэша иначе. В TS-файле модуля маршрутизации передайте объект конфигурации в функцию `forRoot`, которая определяет стратегию. Ниже приведен код в качестве примера. 

```js
import { RouterModule, Routes } from '@angular/router';
// Other imports suppressed for brevity

const routes: Routes = // route definitions go here

@NgModule({
  imports: [RouterModule.forRoot(routes, { useHash: true })],
  exports: [RouterModule]
})
export class AppRoutingModule { }
```   


## <a name="consider-wrapping-fabric-components-with-angular-components"></a>Советуем разместить компоненты Fabric в компонентах Angular

Рекомендуем использовать в надстройке стили [Office UI Fabric](https://developer.microsoft.com/fabric#/fabric-js). Fabric включает компоненты, для которых предусмотрено несколько версий, в том числе версии [на основе TypeScript](https://github.com/OfficeDev/office-ui-fabric-js). Советуем использовать компоненты Fabric в надстройке, размещая их в компонентах Angular. Пример см. в статье [Надстройка проверки стиля в программе Word на основе Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker). Обратите внимание, например, как компонент Angular, определенный в [fabric.textfield.wrapper](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/shared/office-fabric-component-wrappers/fabric.textfield.wrapper.component.ts), импортирует файл TextField.ts Fabric, где определен компонент Fabric. 


## <a name="using-the-office-dialog-api-with-angular"></a>Использование API диалоговых окон Office с Angular

Благодаря API диалоговых окон надстройка Office может открывать страницу в полумодальном диалоговом окне, которое обменивается данными с главной страницей (как правило, в области задач). 

В методе [displayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) используется параметр, указывающий URL-адрес страницы, которая должна открываться в диалоговом окне. В вашей надстройке может быть отдельная HTML-страница, отличная от базовой, для передачи этому параметру. Можно также передать URL-адрес маршрута в приложении на основе Angular. 

Важно помнить, что в случае передачи маршрута диалоговое окно создает новое окно с собственным контекстом выполнения. Базовая страница со всем ее кодом инициализации и начальной загрузки запускается снова в этом новом контексте, а возможным переменным присваиваются первоначальные значения в диалоговом окне. Такой способ приводит к запуску второго экземпляра одностраничного приложения в диалоговом окне. Код, меняющий переменные в диалоговом окне, не меняет версию области задач этих переменных. Для диалогового окна предусмотрено отдельное хранилище сеанса, недоступное из кода в области задач.  


## <a name="trigger-the-ui-update"></a>Запуск обновления пользовательского интерфейса

В приложении Angular пользовательский интерфейс иногда не обновляется. Это происходит потому, что эта часть кода выполняется вне зоны Angular. Чтобы решить эту проблему, поместите код в зону, как показано в приведенном ниже примере.

```js
import { NgZone } from '@angular/core';

export class MyComponent {
  constructor(private zone: NgZone) { }

  myFunction() {
    this.zone.run(() => {
      // the codes that need update the UI
    });
  }
}
``` 

## <a name="using-observable"></a>Использование Observable

Angular использует библиотеку RxJS, в которой предусмотрены объекты `Observable` и `Observer` для реализации асинхронной обработки. Из этого раздела вы узнаете, как использовать `Observables`. Более подробную информацию см. в официальной документации по [RxJS](http://reactivex.io/rxjs/).

Объект `Observable` отчасти похож на объект `Promise`: он возвращается сразу же после асинхронного вызова, но для его разрешения может потребоваться некоторое время. Но если `Promise` — это единственное значение (которое может быть объектом массивов), то `Observable` — это массив объектов (возможно, только с одним элементом). Благодаря этому код может вызывать такие [методы массива](https://www.w3schools.com/jsref/jsref_obj_array.asp), как `concat`, `map` и `filter`, для объектов `Observable`. 

### <a name="pushing-instead-of-pulling"></a>Рассылка вместо извлечения

Ваш код извлекает объекты `Promise`, назначая их переменным, тогда как объекты `Observable` рассылают свои значения объектам, которые *подписаны* на `Observable`. Подписчики — объекты `Observer`. Преимущество подхода, предусматривающего подобную рассылку, состоит в том, что позже можно добавлять в массив `Observable` новые элементы. При добавлении нового элемента все объекты `Observer`, подписанные на `Observable`, получают уведомление. 

Объект `Observer` настроен на обработку каждого нового объекта (именуемого "следующим") с помощью функции. (Он также настроен на то, чтобы отвечать на ошибку и уведомление о завершении. См. пример в следующем разделе.) По этой причине объекты `Observable` можно использовать в более широком диапазоне сценариев, чем объекты `Promise`. Например, в дополнение к возврату `Observable` при вызове AJAX (этим способом можно вернуть также `Promise`), объект `Observable` можно возвращать из обработчика событий, например обработчика событий изменения для текстового поля. Каждый раз, когда пользователь вводит текст в поле, все подписанные объекты `Observer` немедленно реагируют, используя последний текст или текущее состояние приложения в качестве вводных данных. 


### <a name="waiting-until-all-asynchronous-calls-have-completed"></a>Ожидание выполнения всех асинхронных вызовов

Чтобы обратный вызов выполнялся только при условии разрешения каждого элемента из набора объектов `Promise`, используйте метод `Promise.all()`.

```js
myPromise.all([x, y, z]).then(
  // TODO: Callback logic goes here
)
``` 

Чтобы сделать то же самое с объектом `Observable`, используйте метод [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md).  

```js
const source = Observable.forkJoin([x, y, z]);

const subscription = source.subscribe(
  x => {
    // TODO: Callback logic goes here
  },
  err => console.log('Error: ' + err),
  () => console.log('Completed')
);
``` 

## <a name="compile-the-angular-application-using-the-ahead-of-time-aot-compiler"></a>Компиляция приложения Angular с помощью компилятора Ahead-of-Time (AOT)

Производительность приложения — одна из наиболее важных составляющих взаимодействия с пользователем. Компилятор Ahead-of-Time (AOT) позволяет оптимизировать приложение Angular, чтобы компилировать приложение во время сборки. Он полностью преобразовывает исходный код (шаблоны HTML и TypeScript) в эффективный код JavaScript. Если приложение скомпилировано с помощью компилятора AOT, в среде выполнения не будет происходить дополнительная компиляция, что ускорит обработку и выполнение асинхронных запросов для шаблонов HTML. Кроме того, уменьшится общий размер приложения, так как компилятор Angular не придется включать в распространяемый файл приложения. 

Чтобы использовать компилятор AOT, добавьте `--aot` к команде `ng build` или `ng serve`:

```bash
ng build --aot
ng serve --aot
```

> [!NOTE]
> Дополнительные сведения о компиляторе Ahead-of-Time (AOT) приложения Angular см. в [официальном руководстве](https://angular.io/guide/aot-compiler).
