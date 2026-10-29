import * as strings from 'ExtDataSourcesStrings';

/**
 * Groups file extensions by application or file kind.
 * "icon" is a representative extension for the Fluent UI file icon.
 */
export interface IFileTypeGroup {
  key: string;
  label: string;
  icon: string;
  extensions: string[];
}

export const FILE_TYPE_GROUPS: IFileTypeGroup[] = [
  { key: 'word', label: 'Word', icon: 'docx', extensions: ['doc', 'docx', 'docm', 'docb', 'dot', 'dotx', 'dotm'] },
  { key: 'excel', label: 'Excel', icon: 'xlsx', extensions: ['xls', 'xlsx', 'xlsm', 'xlsb', 'xlt', 'xltx', 'xltm', 'xlc', 'xlw', 'xlam'] },
  { key: 'powerpoint', label: 'PowerPoint', icon: 'pptx', extensions: ['ppt', 'pptx', 'pptm', 'pps', 'ppsx', 'ppsm', 'pot', 'potx', 'potm', 'sldx', 'sldm'] },
  { key: 'pdf', label: 'PDF', icon: 'pdf', extensions: ['pdf'] },
  { key: 'onenote', label: 'OneNote', icon: 'one', extensions: ['one', 'onepart', 'onepkg', 'onetoc', 'onetoc2'] },
  { key: 'visio', label: 'Visio', icon: 'vsdx', extensions: ['vsd', 'vsdx', 'vsdm', 'vdx', 'vdw', 'vss', 'vssx', 'vssm', 'vst', 'vstx', 'vstm', 'vsw'] },
  { key: 'project', label: 'Project', icon: 'mpp', extensions: ['mpp', 'mpt'] },
  { key: 'access', label: 'Access', icon: 'accdb', extensions: ['accdb', 'accde', 'mdb'] },
  { key: 'publisher', label: 'Publisher', icon: 'pub', extensions: ['pub'] },
  { key: 'email', get label() { return strings.FileGroupEmail; }, icon: 'msg', extensions: ['msg', 'eml', 'oft'] },
  { key: 'csv', label: 'CSV', icon: 'csv', extensions: ['csv'] },
  { key: 'text', get label() { return strings.FileGroupText; }, icon: 'txt', extensions: ['txt', 'text', 'log', 'md', 'rtf'] },
  { key: 'image', get label() { return strings.FileGroupImages; }, icon: 'jpg', extensions: ['jpg', 'jpeg', 'jpe', 'png', 'gif', 'bmp', 'tif', 'tiff', 'heic', 'heif', 'webp', 'svg', 'ico', 'psd', 'raw'] },
  { key: 'video', get label() { return strings.FileGroupVideos; }, icon: 'mp4', extensions: ['mp4', 'm4v', 'mov', 'avi', 'wmv', 'mkv', 'webm', 'mpg', 'mpeg', 'flv', '3gp'] },
  { key: 'audio', get label() { return strings.FileGroupAudio; }, icon: 'mp3', extensions: ['mp3', 'wav', 'm4a', 'wma', 'aac', 'flac', 'ogg', 'aif', 'aiff'] },
  { key: 'archive', get label() { return strings.FileGroupArchives; }, icon: 'zip', extensions: ['zip', '7z', 'rar', 'gz', 'tgz', 'tar', 'iso', 'cab'] },
  { key: 'sharepoint', get label() { return strings.FileGroupSharePointPages; }, icon: 'aspx', extensions: ['aspx'] },
  { key: 'web', get label() { return strings.FileGroupWebPages; }, icon: 'html', extensions: ['html', 'htm', 'mht', 'mhtml'] },
  { key: 'loop', label: 'Loop', icon: 'loop', extensions: ['loop', 'fluid'] },
  { key: 'whiteboard', label: 'Whiteboard', icon: 'whiteboard', extensions: ['whiteboard', 'wbtx'] },
  { key: 'cad', get label() { return strings.FileGroupCad; }, icon: 'dwg', extensions: ['dwg', 'dxf', 'dwfx', 'step', 'stp', 'stl', 'obj', 'fbx', 'glb', 'gltf', '3mf', 'skp'] },
  { key: 'xml', get label() { return strings.FileGroupXml; }, icon: 'xml', extensions: ['xml', 'xsl', 'xslt', 'json', 'yaml', 'yml'] }
];

const EXTENSION_INDEX: { [ext: string]: IFileTypeGroup } = {};
FILE_TYPE_GROUPS.forEach(g => g.extensions.forEach(e => { EXTENSION_INDEX[e] = g; }));

/** Normalizes a refiner name (".XLSX", "xlsx") to an extension */
export function normalizeExtension(value: string): string {
  return String(value || '').trim().toLowerCase().replace(/^\./, '');
}

/** Group of an extension; unknown extensions form their own group */
export function getFileTypeGroup(value: string): IFileTypeGroup {
  const ext: string = normalizeExtension(value);
  return EXTENSION_INDEX[ext] || { key: `ext:${ext}`, label: ext ? ext.toUpperCase() : strings.FileGroupNoType, icon: ext || 'genericfile', extensions: [ext] };
}
