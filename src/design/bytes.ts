import { sha256 } from "@noble/hashes/sha256";
import { bytesToHex } from "@noble/hashes/utils";

// Pure JS works on the product's LAN HTTP origin, where crypto.subtle is absent.
export const sha256Hex = (bytes: Uint8Array): string => bytesToHex(sha256(bytes));
export function base64ToBytes(value: string): Uint8Array {
  const decoded = atob(value), bytes = new Uint8Array(decoded.length);
  for (let i = 0; i < decoded.length; i++) bytes[i] = decoded.charCodeAt(i);
  return bytes;
}
export function bytesToBase64(bytes: Uint8Array): string {
  let text = "";
  for (let i = 0; i < bytes.length; i += 32768) text += String.fromCharCode(...bytes.subarray(i, i + 32768));
  return btoa(text);
}
