import * as fs from "fs";

export const checkDirectory = async (path) => {
  try {
    await fs.promises.access(path);
    return true;
  } catch (error) {
    return false;
  }
};
