import { MongoClient } from "mongodb";

let clientPromise;

function withTimeout(promise, timeoutMs) {
  return Promise.race([
    promise,
    new Promise((_, reject) => {
      setTimeout(() => reject(new Error("MongoDB connection timed out. Check Atlas Network Access, database password, and internet connection.")), timeoutMs);
    })
  ]);
}

function getMongoUri() {
  const uri = process.env.MONGODB_URI;

  if (!uri || uri.includes("<db_password>")) {
    throw new Error("MongoDB is not configured. Set MONGODB_URI in .env with your real database password.");
  }

  if (uri.startsWith("mongodb+srv://") && uri.includes("cluster0.quz9ta8.mongodb.net")) {
    return expandKnownAtlasSrvUri(uri);
  }

  return uri;
}

function expandKnownAtlasSrvUri(uri) {
  const match = uri.match(/^mongodb\+srv:\/\/([^@]+)@cluster0\.quz9ta8\.mongodb\.net\/?\??(.*)$/);

  if (!match) {
    return uri;
  }

  const credentials = match[1];
  const params = new URLSearchParams(match[2] || "");
  params.set("tls", "true");
  params.set("replicaSet", "atlas-tznd8l-shard-0");
  params.set("authSource", "admin");
  params.set("retryWrites", "true");
  params.set("w", "majority");

  const hosts = [
    "ac-4uwk7eh-shard-00-00.quz9ta8.mongodb.net:27017",
    "ac-4uwk7eh-shard-00-01.quz9ta8.mongodb.net:27017",
    "ac-4uwk7eh-shard-00-02.quz9ta8.mongodb.net:27017"
  ].join(",");

  return `mongodb://${credentials}@${hosts}/?${params.toString()}`;
}

export async function getDb() {
  if (!clientPromise) {
    const client = new MongoClient(getMongoUri(), {
      serverSelectionTimeoutMS: 5000,
      connectTimeoutMS: 5000,
      socketTimeoutMS: 5000
    });
    clientPromise = withTimeout(client.connect(), 6000).catch((error) => {
      clientPromise = undefined;
      throw error;
    });
  }

  const client = await clientPromise;
  return client.db(process.env.MONGODB_DB || "worksight");
}
