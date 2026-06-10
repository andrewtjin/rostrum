// Lets `tsc --noEmit` accept the entry-point stylesheet imports (webpack handles the
// actual loading via css-loader + mini-css-extract-plugin; ts-loader runs transpileOnly).
declare module "*.css";
