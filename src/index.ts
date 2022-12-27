import { init, createOnOpen, askExecute, main } from "./main";

/**
 * @file GASエディタから実行できる関数を定義する
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
declare const global: any;
global.init = init;
global.createOnOpen = createOnOpen;
global.askExecute = askExecute;
global.main = main;
