export interface ProcessResult {
  success: boolean;
  error?: string;
  status?: string;
  counts?: { uploaded: number; skipped: number; renamed: number };
}

export interface TeamsUploadResponse {
  id: string;
}

export interface TeamsMetadataResponse {
  success: boolean;
}

export interface FileDetails {
  name: string;
  size: number;
  type: string;
  base64: string;
}

export interface TeamsTokens {
  chatsvcagg: string | null;
  ic3: string | null;
  permissionsId: string | null;
}

export interface EmojiPreview {
  documentId: string;
  names: string[];
  image?: string;
  error?: string;
  loading?: boolean;
}

export interface EmojiLibraryPage {
  success: true;
  tenantId: string;
  total: number;
  matched: number;
  page: number;
  pageSize: number;
  items: EmojiPreview[];
}

export interface EmojiDuplicates {
  success: true;
  tenantId: string;
  total: number;
  checked: number;
  groups: {
    kind: "name" | "image";
    value: string;
    items: { documentId: string; names: string[] }[];
  }[];
  errors: { names: string[]; error: string }[];
}
