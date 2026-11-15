using System;
using System.Collections.Generic;
using JameelCopilot.Services.Prefs;
using Xunit;

namespace JameelCopilot.Tests.Services.Prefs
{
    public class UserPrefsServiceTests
    {
        private sealed class FakeStore : IUserStore
        {
            public readonly Dictionary<string, string> Items = new Dictionary<string, string>();
            public int Loads;
            public bool FailSaves;
            public bool ThrowOnLoad;
            public bool ThrowOnSave;

            public string Load(string name)
            {
                Loads++;
                if (ThrowOnLoad) throw new InvalidOperationException("store unavailable");
                return Items.TryGetValue(name, out var json) ? json : null;
            }

            public bool Save(string name, string json)
            {
                if (ThrowOnSave) throw new InvalidOperationException("store unavailable");
                if (FailSaves) return false;
                Items[name] = json;
                return true;
            }

            public bool Delete(string name)
            {
                return Items.Remove(name);
            }
        }

        private static FakeStore StoreWith(string json)
        {
            var store = new FakeStore();
            store.Items["Preferences"] = json;
            return store;
        }

        [Fact]
        public void NothingStored_IsBalanced()
        {
            Assert.Equal("balanced", new UserPrefsService(new FakeStore()).Style);
        }

        [Fact]
        public void StoredStyle_IsLoadedOnce()
        {
            var store = StoreWith("{\"v\":1,\"style\":\"creative\"}");
            var prefs = new UserPrefsService(store);

            Assert.Equal("creative", prefs.Style);
            Assert.Equal("creative", prefs.Style);
            Assert.Equal(1, store.Loads);
        }

        [Theory]
        [InlineData("not json")]
        [InlineData("{\"style\":5}")]
        [InlineData("{\"style\":\"loud\"}")]
        [InlineData("{\"style\":{\"a\":1}}")]
        [InlineData("[]")]
        [InlineData("")]
        public void BadStoredData_IsBalanced(string stored)
        {
            Assert.Equal("balanced", new UserPrefsService(StoreWith(stored)).Style);
        }

        [Fact]
        public void SetStyle_SavesVersionAndStyle()
        {
            var store = new FakeStore();
            var prefs = new UserPrefsService(store);

            prefs.SetStyle("creative");

            Assert.Equal("{\"v\":1,\"style\":\"creative\"}", store.Items["Preferences"]);
            Assert.Equal("creative", prefs.Style);
            Assert.Equal("creative", new UserPrefsService(store).Style);
        }

        [Fact]
        public void SetStyle_Unknown_SavesBalanced()
        {
            var store = StoreWith("{\"v\":1,\"style\":\"creative\"}");
            var prefs = new UserPrefsService(store);

            prefs.SetStyle("loud");

            Assert.Equal("{\"v\":1,\"style\":\"balanced\"}", store.Items["Preferences"]);
            Assert.Equal("balanced", prefs.Style);
        }

        // Review Focus 3: a mailbox that cannot keep the item (shared, PST-only, offline).
        [Theory]
        [InlineData(true, false)]
        [InlineData(false, true)]
        public void SaveFailsOrThrows_TheChoiceStillHoldsForTheSession(bool failSaves, bool throwOnSave)
        {
            var prefs = new UserPrefsService(new FakeStore { FailSaves = failSaves, ThrowOnSave = throwOnSave });

            prefs.SetStyle("precise");

            Assert.Equal("precise", prefs.Style);
        }

        [Fact]
        public void StoreThrowsOnLoad_IsBalanced_AndDoesNotThrow()
        {
            var prefs = new UserPrefsService(new FakeStore { ThrowOnLoad = true });

            Assert.Equal("balanced", prefs.Style);
            prefs.SetStyle("creative");
            Assert.Equal("creative", prefs.Style);
        }

        [Fact]
        public void NullStore_KeepsTheChoiceInMemory()
        {
            var prefs = new UserPrefsService(null);

            Assert.Equal("balanced", prefs.Style);
            prefs.SetStyle("creative");
            Assert.Equal("creative", prefs.Style);
        }
    }
}
