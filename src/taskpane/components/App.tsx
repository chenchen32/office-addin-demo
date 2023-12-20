import React from "react";
import { Button } from "antd";
import * as officeUtils from "../office-document";

const App = () => {
  const insertText = () => {
    officeUtils.insertText("一些文字。。。");
  };

  const insertImage = () => {
    officeUtils.insertImage("https://cdn.jsdelivr.net/gh/chenchen32/office-addin-demo@main/public/assets/squirtle.jpg");
  };

  const insertPPT = (orderNum: number) => {
    officeUtils.applyPPTByIndex(
      "https://raw.githubusercontent.com/chenchen32/office-addin-demo/main/public/assets/sport.pptx",
      orderNum - 1
    );
  };

  return (
    <div style={{ textAlign: "center" }}>
      <div className="button-section">
        <Button type="primary" onClick={insertText}>
          插入文字
        </Button>
      </div>
      <div className="button-section">
        <Button type="primary" onClick={insertImage}>
          插入图片
        </Button>
      </div>
      {[1, 2, 3].map((orderNum) => {
        return (
          <div key={orderNum} className="button-section">
            <Button type="primary" onClick={() => insertPPT(orderNum)}>
              插入PPT第{orderNum}页
            </Button>
          </div>
        );
      })}
    </div>
  );
};

export default App;
