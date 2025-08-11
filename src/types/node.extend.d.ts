declare module 'node.extend' {
  type AnyRecord = Record<string, any>;

  /**
   * Extends the target object with the provided source objects.
   * When `deep` is true, performs a deep (recursive) merge.
   */
  export default function extend<TTarget extends AnyRecord, TSource extends AnyRecord>(
    deep: boolean,
    target: TTarget,
    ...sources: TSource[]
  ): TTarget & TSource;
}


