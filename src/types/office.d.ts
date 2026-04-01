/* eslint-disable @typescript-eslint/no-explicit-any */
/**
 * Office.js 全局类型声明
 * 补充 @types/office-js 可能缺失的类型
 */

declare namespace PowerPoint {
  function run(
    callback: (context: PowerPoint.RequestContext) => Promise<void>
  ): Promise<void>;

  interface RequestContext {
    presentation: Presentation;
    sync(): Promise<void>;
  }

  interface Presentation {
    slides: SlideCollection;
    getSelectedSlides(): SlideCollection;
    insertSlidesFromBase64(
      base64: string,
      options?: InsertSlideOptions
    ): void;
  }

  interface InsertSlideOptions {
    formatting?: string;
    targetSlideId?: string;
    sourceSlideIds?: string[];
  }

  interface SlideCollection {
    items: Slide[];
    load(properties: string): void;
    getItemAt(index: number): Slide;
  }

  interface Slide {
    id: string;
    shapes: ShapeCollection;
  }

  interface ShapeCollection {
    addImage(base64: string): Shape;
  }

  interface Shape {
    id: string;
    width: number;
    height: number;
    left: number;
    top: number;
  }
}

declare namespace Office {
  interface SetSelectedDataOptions {
    coercionType: any;
    imageLeft?: number;
    imageTop?: number;
    imageWidth?: number;
    imageHeight?: number;
  }
}
