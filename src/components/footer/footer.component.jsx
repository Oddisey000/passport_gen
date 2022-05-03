import React from "react";
import "./footer.component.scss";

const FooterGrid = () => {
  return (
    <span className="footer-text splashInBottom">
      This application is made by{" "}
      <a
        className="link"
        href="https://www.facebook.com/oddisey000"
      >
        Vitalii Pertsovych
      </a>{" "}
      in Kolomyia.
    </span>
  );
};

export default FooterGrid;
