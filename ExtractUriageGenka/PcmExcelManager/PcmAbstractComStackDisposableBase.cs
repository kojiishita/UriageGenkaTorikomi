
namespace PCM
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;

    /// <summary>
    /// COM <see cref="Stack{T}"/> 破棄可能抽象基底クラスです。
    /// </summary>
    /// <remarks>
    /// COM にアクセスするクラスの継承元となるクラスです。
    /// アクセスした COM オブジェクトを <see cref="Stack"/> にスタックします。
    /// <see cref="Dispose"/> にて COM オブジェクトを開放します。
    /// </remarks>
    public abstract class PcmAbstractComStackDisposableBase : IDisposable
    {
        /// <summary><see cref="Stack{T}"/> オブジェクトです。</summary>
        protected Stack<object> Stack { get; set; } = new Stack<object>();

        /// <summary>
        /// リソースを開放します。
        /// </summary>
        public virtual void Dispose()
        {
            this.Release();

            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// <see cref="Stack"/> を開放します。
        /// </summary>
        protected void Release()
        {
            while (this.Stack.Count > 0)
            {
                Marshal.ReleaseComObject(this.Stack.Pop());
            }

            this.Stack.Clear();
        }
    }
}