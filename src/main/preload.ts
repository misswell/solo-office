import { contextBridge, ipcRenderer } from "electron";

contextBridge.exposeInMainWorld("soloOffice", {
  pickFile: (): Promise<{ name: string; path: string; data: Uint8Array } | null> => ipcRenderer.invoke("solo:pick"),
  pickImage: (): Promise<{ name: string; data: Uint8Array } | null> => ipcRenderer.invoke("solo:pick-image"),
  saveFile: (name: string, data: Uint8Array | string): Promise<string | null> => ipcRenderer.invoke("solo:save", { name, data }),
  writeFile: (path: string, data: Uint8Array): Promise<string | null> => ipcRenderer.invoke("solo:write", { path, data }),
  confirmUnsaved: (name: string): Promise<"save" | "discard" | "cancel"> => ipcRenderer.invoke("solo:confirm", name),
  closeReply: (proceed: boolean): void => ipcRenderer.send("solo:close-reply", proceed),
  setFullscreen: (on: boolean): void => ipcRenderer.send("solo:fullscreen", on),
  exportPdf: (name: string, landscape: boolean): Promise<string | null> => ipcRenderer.invoke("solo:pdf", { name, landscape }),
});

const relay = (channel: string): void => {
  ipcRenderer.on(channel, (_e, detail) => {
    window.dispatchEvent(new CustomEvent(channel, { detail }));
  });
};
for (const channel of ["solo:load", "solo:open", "solo:save", "solo:saveas", "solo:new", "solo:show", "solo:export-pdf", "solo:closerequest"]) relay(channel);
