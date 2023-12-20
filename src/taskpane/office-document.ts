/* global Office PowerPoint console */
import JSZip from "jszip";
import { requestPure } from "./api/request";

const getBase64 = (arrayBuffer) => {
  const uInt8Array = new Uint8Array(arrayBuffer);
  let i = uInt8Array.length;
  const binaryString = new Array(i);
  while (i--) {
    binaryString[i] = String.fromCharCode(uInt8Array[i]);
  }
  const data = binaryString.join("");
  // eslint-disable-next-line no-undef
  return window.btoa(data);
};

export const insertText = async (text: string) => {
  try {
    Office.context.document.setSelectedDataAsync(
      text,
      {
        coercionType: Office.CoercionType.Text,
      },
      (result: Office.AsyncResult<void>) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          throw result.error.message;
        }
      }
    );
  } catch (error) {
    console.error("insert text error: ", error);
  }
};

export const insertImage = async (imageUrl: string) => {
  try {
    const res = await requestPure({
      url: imageUrl,
      method: "GET",
      responseType: "arraybuffer",
    });

    const base64 = getBase64(res);
    Office.context.document.setSelectedDataAsync(
      base64,
      {
        coercionType: Office.CoercionType.Image,
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          throw result.error.message;
        }
      }
    );
  } catch (error) {
    console.error("insert image error: ", error);
  }
};

export const getActiveSlideId = () => {
  return new Promise((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        // @ts-expect-error
        const slideId = result.value.slides[0].id;
        resolve(slideId);
      } else {
        reject(`获取 PowerPoint 激活页面的 id 失败：${result.error.message}`);
      }
    });
  });
};

export const insertPPT = (pptUrl: string, options: PowerPoint.InsertSlideOptions) => {
  return PowerPoint.run(async (context) => {
    try {
      const res = await requestPure({
        url: pptUrl,
        method: "GET",
        responseType: "arraybuffer",
      });

      const base64 = getBase64(res);

      context.presentation.insertSlidesFromBase64(base64, options);
    } catch (error) {
      console.error("insert ppt error: ", error);
    }
  });
};

export const getPPTManifest = async (url: string) => {
  const inputFile = await requestPure({
    url,
    method: "GET",
    responseType: "arraybuffer",
  });

  const pptZip = await JSZip().loadAsync(inputFile);
  const json = {};
  await Promise.all(
    Object.keys(pptZip.files).map(async (relativePath) => {
      const file = pptZip.file(relativePath);
      const ext = relativePath.split(".").pop();

      let content;
      if (!file || file.dir) {
        return;
      } else if (["xml", "rels"].includes(ext)) {
        const xml = await file.async("string");
        const parser = new DOMParser();
        content = parser.parseFromString(xml, "text/xml");
      } else {
        // images, audio files, movies, etc.
        content = await file.async("arraybuffer");
      }
      json[relativePath] = content;
    })
  );
  return json;
};

export const getSlideIdsFromManifest = (json: Record<string, any>) => {
  const presentation = json["ppt/presentation.xml"] as Document;
  const sldIdLst = presentation.getElementsByTagName("p:sldIdLst")[0];
  const ids = [];
  sldIdLst.childNodes.forEach((node) => {
    // @ts-expect-error
    ids.push(node.id);
  });
  return ids;
};

export const applyPPTByIndex = async (pptUrl: string, index: number) => {
  // 获取当前激活的页面 id，从当前页面插入
  const targetSlideId = await getActiveSlideId();
  // 获取 ppt 文件解压缩之后的清单信息
  const json = await getPPTManifest(pptUrl);
  // 获取 ppt 所有页面的 id list
  const ids = getSlideIdsFromManifest(json);

  // 插入某一页 ppt
  return insertPPT(pptUrl, {
    targetSlideId: `${targetSlideId}`,
    sourceSlideIds: [ids[index]],
  });
};
