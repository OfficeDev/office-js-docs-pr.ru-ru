---
title: Разработка надстроек Office с помощью Angular
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 65b2a229e0379106b63b0f1abaaa8b66d7cdf367
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/25/2018
ms.locfileid: "25004976"
---
# <a name="develop-office-add-ins-with-angular"></a><span data-ttu-id="03e54-102">Разработка надстроек Office с помощью Angular</span><span class="sxs-lookup"><span data-stu-id="03e54-102">Develop Office Add-ins with Angular</span></span>

<span data-ttu-id="03e54-103">В этой статье приведены рекомендации по использованию Angular 2 и более поздних версий для создания надстройки Office в виде одностраничного приложения.</span><span class="sxs-lookup"><span data-stu-id="03e54-103">This article provides guidance for using Angular 2+ to create an Office Add-in as a single page application.</span></span>

> [!NOTE]
> <span data-ttu-id="03e54-p101">Вы можете поделиться опытом по использованию Angular для создания надстроек Office? Примите участие в создании этой статьи на сайте [GitHub](https://github.com/OfficeDev/office-js-docs) или сообщите о [проблеме](https://github.com/OfficeDev/office-js-docs-pr/issues) в соответствующем репозитории.</span><span class="sxs-lookup"><span data-stu-id="03e54-p101">Do you have something to contribute based on your experience using Angular to create Office Add-ins? You can contribute to this article in [GitHub](https://github.com/OfficeDev/office-js-docs) or provide your feedback by submitting an [issue](https://github.com/OfficeDev/office-js-docs-pr/issues) in the repo.</span></span> 

<span data-ttu-id="03e54-106">Пример надстройки Office, созданной на платформе Angular, приведен в статье [Надстройка на основе Angular для проверки стиля в приложении Word](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).</span><span class="sxs-lookup"><span data-stu-id="03e54-106">For an Office Add-ins sample that's built using the Angular framework, see [Word Style Checking Add-in Built on Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).</span></span>

## <a name="install-the-typescript-type-definitions"></a><span data-ttu-id="03e54-107">Установка определений типов TypeScript</span><span class="sxs-lookup"><span data-stu-id="03e54-107">Install the TypeScript type definitions</span></span>
<span data-ttu-id="03e54-108">Откройте окно nodejs и введите в командной строке следующую команду:</span><span class="sxs-lookup"><span data-stu-id="03e54-108">Open an nodejs window and enter the following at the command line:</span></span> 

```bash
npm install --save-dev @types/office-js
```

## <a name="bootstrapping-must-be-inside-officeinitialize"></a><span data-ttu-id="03e54-109">Начальная загрузка должна определяться в методе Office.initialize</span><span class="sxs-lookup"><span data-stu-id="03e54-109">Bootstrapping must be inside Office.initialize</span></span>

<span data-ttu-id="03e54-p102">На любой странице, которая вызывает интерфейсы API JavaScript для Office, Word или Excel, в коде сначала нужно назначить метод для свойства `Office.initialize`. (Если у вас нет кода инициализации, тело метода может состоять из пустых символов "`{}`", но свойство `Office.initialize` должно быть определено. Дополнительные сведения см. в разделе [Инициализация надстройки](understanding-the-javascript-api-for-office.md#initializing-your-add-in).) Office вызывает этот метод сразу же после того, как инициализирует библиотеки JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="03e54-p102">On any page that calls the Office, Word, or Excel JavaScript APIs, your code must first assign a method to the `Office.initialize` property. (If you have no initialization code, the method body can be just empty "`{}`" symbols, but you must not leave the `Office.initialize` property undefined. For details, see [Initializing your add-in](understanding-the-javascript-api-for-office.md#initializing-your-add-in).) Office calls this method immediately after it has initialized the Office JavaScript libraries.</span></span>

<span data-ttu-id="03e54-p103">**Вызов кода начальной загрузки на основе Angular необходимо задать в методе, который назначен `Office.initialize`**, чтобы сначала выполнялась инициализация библиотек JavaScript для Office. Вот простой пример, в котором показано, как это сделать. Этот код должен находиться в файле main.ts проекта.</span><span class="sxs-lookup"><span data-stu-id="03e54-p103">**Your Angular bootstrapping code must be called inside the method that you assign to `Office.initialize`** to ensure that the Office JavaScript libraries have initialized first. The following is a simple example that shows how to do this. This code should be in the main.ts file of the project.</span></span>

```js
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { AppModule } from './app.module';

Office.initialize = function () {
  const platform = platformBrowserDynamic();
  platform.bootstrapModule(AppModule);
};
```

## <a name="use-the-hash-location-strategy-in-the-angular-application"></a><span data-ttu-id="03e54-116">Используйте стратегию навигации с помощью хэша в приложении на основе Angular</span><span class="sxs-lookup"><span data-stu-id="03e54-116">Use the hash location strategy in the Angular application</span></span>

<span data-ttu-id="03e54-p104">Переход между маршрутами в приложении может не выполняться, если не задать стратегию навигации с помощью хэша. Это можно сделать одним из двух способов. Способ первый: указать поставщика стратегии навигации в модуле приложения, как показано в приведенном ниже примере. (Это для файла app.module.ts.)</span><span class="sxs-lookup"><span data-stu-id="03e54-p104">Navigating between routes in the application might not work if you don't specify the hash location strategy. You can do this in one of two ways. First, you can specify a provider for the location strategy in your app module, as shown in the following example. It goes into the app.module.ts file.</span></span>

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

<span data-ttu-id="03e54-p105">Если маршруты определены в отдельном модуле маршрутизации, можно задать стратегию навигации с помощью хэша иначе. В TS-файле модуля маршрутизации передайте объект конфигурации в функцию `forRoot`, которая определяет стратегию. Ниже приведен код в качестве примера.</span><span class="sxs-lookup"><span data-stu-id="03e54-p105">If you define your routes in a separate routing module, there is an alternative way to specify the hash location strategy. In your routing module's .ts file, pass a configuration object to the `forRoot` function that specifies the strategy. The following code is an example.</span></span> 

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


## <a name="consider-wrapping-fabric-components-with-angular-components"></a><span data-ttu-id="03e54-124">Советуем разместить компоненты Fabric в компонентах Angular</span><span class="sxs-lookup"><span data-stu-id="03e54-124">Consider wrapping Fabric components with Angular components</span></span>

<span data-ttu-id="03e54-p106">Рекомендуем использовать в надстройке стили [Office UI Fabric](https://developer.microsoft.com/fabric#/fabric-js). Fabric включает компоненты, для которых предусмотрено несколько версий, в том числе версии [на основе TypeScript](https://github.com/OfficeDev/office-ui-fabric-js). Советуем использовать компоненты Fabric в надстройке, размещая их в компонентах Angular. Пример см. в статье [Надстройка проверки стиля в программе Word на основе Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker). Обратите внимание, например, как компонент Angular, определенный в [fabric.textfield.wrapper](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/shared/office-fabric-component-wrappers/fabric.textfield.wrapper.component.ts), импортирует файл TextField.ts Fabric, где определен компонент Fabric.</span><span class="sxs-lookup"><span data-stu-id="03e54-p106">We recommend using [Office UI Fabric](https://developer.microsoft.com/fabric#/fabric-js) styling in your add-in. Fabric includes components that come in several versions, including a version [based on TypeScript](https://github.com/OfficeDev/office-ui-fabric-js). Consider using Fabric components in your add-in by wrapping them in Angular components. For an example that shows you how to do this, see [Word Style Checking Add-in Built on Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker). Note, for example, how the Angular component defined in [fabric.textfield.wrapper](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/shared/office-fabric-component-wrappers/fabric.textfield.wrapper.component.ts) imports the Fabric file TextField.ts, where the Fabric component is defined.</span></span> 


## <a name="using-the-office-dialog-api-with-angular"></a><span data-ttu-id="03e54-130">Использование API диалоговых окон Office с Angular</span><span class="sxs-lookup"><span data-stu-id="03e54-130">Using the Office Dialog API with Angular</span></span>

<span data-ttu-id="03e54-131">API диалогового окна надстройки Office позволяет открывать страницу в полумодальном диалоговом окне, способном обменивается данными с главной страницей, которая, как правило, располагается в области задач).</span><span class="sxs-lookup"><span data-stu-id="03e54-131">The Office add-in Dialog API enables your add-in to open a page in a semimodal dialog box that can exchange information with the main page, which is typically in a task pane.</span></span> 

<span data-ttu-id="03e54-p107">В методе [displayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) используется параметр, указывающий URL-адрес страницы, которая должна открываться в диалоговом окне. В вашей надстройке может быть отдельная HTML-страница, отличная от базовой, для передачи этому параметру. Можно также передать URL-адрес маршрута в приложении на основе Angular.</span><span class="sxs-lookup"><span data-stu-id="03e54-p107">The [displayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) method takes a parameter that specifies the URL of the page that should open in the dialog box. Your add-in can have a separate HTML page (different from the base page) to pass to this parameter, or you can pass the URL of a route in your Angular appication.</span></span> 

<span data-ttu-id="03e54-p108">Важно помнить, что в случае передачи маршрута диалоговое окно создает новое окно с собственным контекстом выполнения. Базовая страница со всем ее кодом инициализации и начальной загрузки запускается снова в этом новом контексте, а возможным переменным присваиваются первоначальные значения в диалоговом окне. Такой способ приводит к запуску второго экземпляра одностраничного приложения в диалоговом окне. Код, меняющий переменные в диалоговом окне, не меняет версию области задач этих переменных. Для диалогового окна предусмотрено отдельное хранилище сеанса, недоступное из кода в области задач.</span><span class="sxs-lookup"><span data-stu-id="03e54-p108">It is important to remember, if you pass a route, that the dialog box creates a new window with its own execution context. Your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog box. So this technique launches a second instance of your single page application in the dialog box. Code that changes variables in the dialog box does not change the task pane version of the same variables. Similarly, the dialog box has its own session storage, which is not accessible from code in the task pane.</span></span>  


## <a name="trigger-the-ui-update"></a><span data-ttu-id="03e54-139">Запуск обновления пользовательского интерфейса</span><span class="sxs-lookup"><span data-stu-id="03e54-139">Trigger the UI update</span></span>

<span data-ttu-id="03e54-p109">В приложении Angular пользовательский интерфейс иногда не обновляется. Это происходит потому, что эта часть кода выполняется вне зоны Angular. Чтобы решить эту проблему, поместите код в зону, как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="03e54-p109">In an Angular app, the UI sometimes does not update. This is because that part of the code runs out of the Angular zone. The solution is to put the code in the zone, as shown in the following example.</span></span>

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

## <a name="using-observable"></a><span data-ttu-id="03e54-143">Использование Observable</span><span class="sxs-lookup"><span data-stu-id="03e54-143">Using Observable</span></span>

<span data-ttu-id="03e54-p110">Angular использует библиотеку RxJS, в которой предусмотрены объекты `Observable` и `Observer` для реализации асинхронной обработки. Из этого раздела вы узнаете, как использовать `Observables`. Более подробную информацию см. в официальной документации по [RxJS](http://reactivex.io/rxjs/).</span><span class="sxs-lookup"><span data-stu-id="03e54-p110">Angular uses RxJS (Reactive Extensions for JavaScript), and RxJS introduces `Observable` and `Observer` objects to implement asynchronous processing. This section provides a brief introduction to using `Observables`; for more detailed information, see the official [RxJS](http://reactivex.io/rxjs/) documentation.</span></span>

<span data-ttu-id="03e54-p111">Объект `Observable` отчасти похож на объект `Promise`: он возвращается сразу же после асинхронного вызова, но для его разрешения может потребоваться некоторое время. Но если `Promise` — это единственное значение (которое может быть объектом массивов), то `Observable` — это массив объектов (возможно, только с одним элементом). Благодаря этому код может вызывать такие [методы массива](https://www.w3schools.com/jsref/jsref_obj_array.asp), как `concat`, `map` и `filter`, для объектов `Observable`.</span><span class="sxs-lookup"><span data-stu-id="03e54-p111">An `Observable` is like a `Promise` object in some ways - it is returned immediately from an asynchronous call, but it might not resolve until some time later. However, while a `Promise` is a single value (which can be an array object), an `Observable` is an array of objects (possibly with only a single member). This enables code to call [array methods](https://www.w3schools.com/jsref/jsref_obj_array.asp), such as `concat`, `map`, and `filter`, on `Observable` objects.</span></span> 

### <a name="pushing-instead-of-pulling"></a><span data-ttu-id="03e54-149">Рассылка вместо извлечения</span><span class="sxs-lookup"><span data-stu-id="03e54-149">Pushing instead of pulling</span></span>

<span data-ttu-id="03e54-p112">Ваш код извлекает объекты `Promise`, назначая их переменным, тогда как объекты `Observable` рассылают свои значения объектам, которые *подписаны* на `Observable`. Подписчики — объекты `Observer`. Преимущество подхода, предусматривающего подобную рассылку, состоит в том, что позже можно добавлять в массив `Observable` новые элементы. При добавлении нового элемента все объекты `Observer`, подписанные на `Observable`, получают уведомление.</span><span class="sxs-lookup"><span data-stu-id="03e54-p112">Your code "pulls" `Promise` objects by assigning them to variables, but `Observable` objects "push" their values to objects that *subscribe* to the `Observable`. The subscribers are `Observer` objects. The benefit of the push architecture is that new members can be added to the `Observable` array over time. When a new member is added, all the `Observer` objects that subscribe to the `Observable` receive a notification.</span></span> 

<span data-ttu-id="03e54-p113">Объект `Observer` настроен на обработку каждого нового объекта (именуемого "следующим") с помощью функции. (Он также настроен на то, чтобы отвечать на ошибку и уведомление о завершении. См. пример в следующем разделе.) По этой причине объекты `Observable` можно использовать в более широком диапазоне сценариев, чем объекты `Promise`. Например, в дополнение к возврату `Observable` при вызове AJAX (этим способом можно вернуть также `Promise`), объект `Observable` можно возвращать из обработчика событий, например обработчика событий изменения для текстового поля. Каждый раз, когда пользователь вводит текст в поле, все подписанные объекты `Observer` немедленно реагируют, используя последний текст или текущее состояние приложения в качестве вводных данных.</span><span class="sxs-lookup"><span data-stu-id="03e54-p113">The `Observer` is configured to process each new object (called the "next" object) with a function. (It is also configured to respond to an error and a completion notification. See the next section for an example.) For this reason, `Observable` objects can be used in a wider range of scenarios than `Promise` objects. For example, in addition to returning an `Observable` from an AJAX call, the way you can return a `Promise`, an `Observable` can be returned from an event handler, such as the "changed" event handler for a text box. Each time a user enters text in the box, all the subscribed `Observer` objects react immediately using the latest text and/or the current state of the application as input.</span></span> 


### <a name="waiting-until-all-asynchronous-calls-have-completed"></a><span data-ttu-id="03e54-159">Ожидание выполнения всех асинхронных вызовов</span><span class="sxs-lookup"><span data-stu-id="03e54-159">Waiting until all asynchronous calls have completed</span></span>

<span data-ttu-id="03e54-160">Чтобы обратный вызов выполнялся только при условии разрешения каждого элемента из набора объектов `Promise`, используйте метод `Promise.all()`.</span><span class="sxs-lookup"><span data-stu-id="03e54-160">When you want to ensure that a callback only runs when every member of a set of `Promise` objects has resolved, use the `Promise.all()` method.</span></span>

```js
myPromise.all([x, y, z]).then(
  // TODO: Callback logic goes here
)
``` 

<span data-ttu-id="03e54-161">Чтобы сделать то же самое с объектом `Observable`, используйте метод [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md).</span><span class="sxs-lookup"><span data-stu-id="03e54-161">To do the same thing with an `Observable` object, you use the [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md) method.</span></span>  

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

