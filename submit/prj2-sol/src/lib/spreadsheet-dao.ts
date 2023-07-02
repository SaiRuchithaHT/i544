import { Result, okResult, errResult } from 'cs544-js-utils';

import * as mongo from 'mongodb';

/** All that this DAO should do is maintain a persistent map from
 *  [spreadsheetName, cellId] to an expression string.
 *
 *  Most routines return an errResult with code set to 'DB' if
 *  a database error occurs.
 */

/** return a DAO for spreadsheet ssName at URL mongodbUrl */
export async function
makeSpreadsheetDao(mongodbUrl: string, ssName: string)
  : Promise<Result<SpreadsheetDao>> 
{
  return SpreadsheetDao.make(mongodbUrl, ssName);
}

export class SpreadsheetDao {
  private spreadsheetName: string;
  private db: mongo.Db;
  private collection: mongo.Collection;
  private client: mongo.MongoClient;

  private constructor(db: mongo.Db, collection: mongo.Collection, ssName: string, client: mongo.MongoClient) {
    this.db = db;
    this.collection = collection;
    this.spreadsheetName = ssName;
    this.client = client;
  }

  //factory method
  static async make(dbUrl: string, ssName: string): Promise<Result<SpreadsheetDao>> {
    try {
      // Connect to the MongoDB server
      const client = await mongo.MongoClient.connect(dbUrl);
      const db = client.db();
      const collection = db.collection(ssName);
      // Create and return a new instance of SpreadsheetDao
      return okResult(new SpreadsheetDao(db, collection, ssName, client));
    } catch (error) {
      return errResult(error.message, 'DB');
    }
  }

  /** Release all resources held by persistent spreadsheet.
   *  Specifically, close any database connections.
   */
  async close(): Promise<Result<undefined>> {
    try {
      // Close the MongoDB client connection
      await this.client.close();
      return okResult(undefined);
    } catch (error) {
      return errResult([{ options: { code: 'DB' }, message: error.message }]);
    }
  }

  /** return name of this spreadsheet */
  getSpreadsheetName(): string {
    return this.spreadsheetName;
  }

  /** Set cell with id cellId to string expr. */
  async setCellExpr(cellId: string, expr: string): Promise<Result<undefined>> {
    try {
      // Update or insert the document with the specified cellId
      await this.collection.updateOne({ _id: cellId }, { $set: { expr } }, { upsert: true });
      return okResult(undefined);
    } catch (error) {
      return errResult([{ options: { code: 'DB' }, message: error.message }]);
    }
  }

  /** Return expr for cell cellId; return '' for an empty/unknown cell.
   */
  async query(cellId: string): Promise<Result<string>> {
    try {
      // Find the document with the specified cellId
      const document = await this.collection.findOne({ _id: cellId });
      if (document) {
        const expr = document.expr;
        return okResult(expr);
      } else {
        return okResult('');
      }
    } catch (error) {
      return errResult([{ options: { code: 'DB' }, message: error.message }]);
    }
  }

  /** Clear contents of this spreadsheet */
  async clear(): Promise<Result<undefined>> {
    try {
      // Delete all documents in the collection
      await this.collection.deleteMany({});
      return okResult(undefined);
    } catch (error) {
      return errResult([{ options: { code: 'DB' }, message: error.message }]);
    }
  }

  /** Remove all info for cellId from this spreadsheet. */
  async remove(cellId: string): Promise<Result<undefined>> {
    try {
      // Delete the document with the specified cellId
      await this.collection.deleteOne({ _id: cellId });
      return okResult(undefined);
    } catch (error) {
      return errResult([{ options: { code: 'DB' }, message: error.message }]);
    }
  }

  /** Return array of [ cellId, expr ] pairs for all cells in this
   *  spreadsheet
   */
  async getData(): Promise<Result<[string, string][]>> {
  try {
    // Retrieve all documents from the collection and map them to [ cellId, expr ] pairs
    const documents = await this.collection.find().toArray();
    const data = documents.map((doc) => [doc._id.toString(), doc.expr] as [string, string]);
    return okResult(data);
  } catch (error) {
    return errResult([{ options: { code: 'DB' }, message: error.message }]);
  }

}
}
