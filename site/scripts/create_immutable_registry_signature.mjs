#!/usr/bin/env node

import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const relativeDirectory = process.argv[2];
const keyId = process.argv[3];

if (!relativeDirectory || !keyId) {
  throw new Error("usage: create_immutable_registry_signature.mjs <registry-directory> <key-id>");
}

const directory = path.resolve(root, relativeDirectory);
assert(
  directory.startsWith(`${path.join(root, "registry")}${path.sep}`),
  "registry directory must be beneath registry/",
);
const indexPath = path.join(directory, "index.json");
const signaturePath = path.join(directory, "index.sig.json");
const publicKeyPath = path.join(directory, "public-key.pem");
assert(fs.existsSync(indexPath), "registry index is missing");
assert(!fs.existsSync(signaturePath), "refusing to replace an immutable registry signature");
assert(!fs.existsSync(publicKeyPath), "refusing to replace an immutable registry public key");

const index = fs.readFileSync(indexPath);
const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
const signature = crypto.sign(null, index, privateKey);
const publicKeyDer = publicKey.export({ type: "spki", format: "der" });
const publicKeyHex = publicKeyDer.subarray(publicKeyDer.length - 32).toString("hex");
const record = {
  schema_version: 1,
  algorithm: "ed25519",
  key_id: keyId,
  manifest_sha256: crypto.createHash("sha256").update(index).digest("hex"),
  signature: signature.toString("hex"),
};

fs.writeFileSync(publicKeyPath, publicKey.export({ type: "spki", format: "pem" }), { flag: "wx" });
fs.writeFileSync(signaturePath, `${JSON.stringify(record, null, 2)}\n`, { flag: "wx" });
assert(crypto.verify(null, index, publicKey, signature), "generated signature did not verify");

console.log(JSON.stringify({
  directory: path.relative(root, directory),
  key_id: keyId,
  public_key_hex: publicKeyHex,
  manifest_sha256: record.manifest_sha256,
}));

function assert(condition, message) {
  if (!condition) throw new Error(message);
}
