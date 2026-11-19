"""Rank-wise GitHub repository URLs.

Local bare repositories only. No GitHub is contacted, no production remote
is used, and the other project's repository is refused by name.
"""
import importlib.util
import inspect
import os
import subprocess
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(HERE, "find_school.py")
_spec = importlib.util.spec_from_file_location("fs", SCRIPT)
fs = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(fs)


def git(where, *args):
    return subprocess.run(["git", "-C", where] + list(args),
                          capture_output=True, text=True)


class RankRepositoryUrlTests(unittest.TestCase):
    def test_every_rank_one_to_ten_gets_its_own_repository(self):
        urls = dict(fs.rank_repo_urls())
        self.assertEqual(sorted(urls), list(range(1, 11)))
        for rank in range(1, 11):
            self.assertTrue(urls[rank].endswith("/school_extractor%d.git" % rank),
                            urls[rank])
        self.assertEqual(len(set(urls.values())), 10, "two ranks share a URL")

    def test_the_rank_goes_before_the_extension_only(self):
        for rank, url in fs.rank_repo_urls():
            self.assertEqual(url, fs.base_github_url[:-len(".git")]
                             + str(rank) + ".git")

    def test_rank_ten_is_not_rank_one_with_a_zero(self):
        urls = dict(fs.rank_repo_urls())
        self.assertNotEqual(urls[10], urls[1] + "0")
        self.assertTrue(urls[10].endswith("school_extractor10.git"))

    def test_the_base_url_stays_configurable(self):
        other = "git@github.com:someone/other_name.git"
        self.assertTrue(fs.rank_repo_url(3, 10, other)
                        .endswith("/other_name3.git"))

    def test_the_effective_rank_comes_from_rank_json(self):
        kept = fs.RANK_FILE
        folder = tempfile.mkdtemp(prefix="rank_json_")
        try:
            fs.RANK_FILE = os.path.join(folder, "rank.json")
            with open(fs.RANK_FILE, "w", encoding="utf-8") as handle:
                handle.write('{"rank": 7, "num_sys": 10}')
            self.assertTrue(fs.rank_repo_url().endswith("school_extractor7.git"))
        finally:
            fs.RANK_FILE = kept

    def test_the_environment_overrides_rank_json(self):
        kept = fs.RANK_FILE
        folder = tempfile.mkdtemp(prefix="rank_env_")
        try:
            fs.RANK_FILE = os.path.join(folder, "rank.json")
            with open(fs.RANK_FILE, "w", encoding="utf-8") as handle:
                handle.write('{"rank": 7, "num_sys": 10}')
            os.environ["SCHOOL_RANK"] = "4"
            self.assertTrue(fs.rank_repo_url().endswith("school_extractor4.git"))
        finally:
            os.environ.pop("SCHOOL_RANK", None)
            fs.RANK_FILE = kept

    def test_a_rank_outside_the_fleet_is_refused(self):
        for bad in (0, 11, -1, "seven"):
            with self.assertRaises(fs.RankRepoError, msg=repr(bad)):
                fs.rank_repo_url(bad, 10)

    def test_no_rank_means_the_effective_rank(self):
        """None is the documented sentinel, not an invalid value: it asks
        for whatever rank this machine resolves to."""
        effective = fs.resolve_rank()[0]
        self.assertEqual(fs.rank_repo_url(None, 10),
                         fs.rank_repo_url(effective, 10))

    def test_a_base_that_is_not_a_repository_url_is_refused(self):
        for bad in ("", "   ", "git@github.com:me/repo", "nonsense"):
            with self.assertRaises(fs.RankRepoError, msg=repr(bad)):
                fs.rank_repo_url(1, 10, bad)

    def test_the_other_projects_repository_can_never_be_a_destination(self):
        forbidden = "git@github.com:evisionscripts/github_profile_extractor.git"
        with self.assertRaises(fs.RankRepoError):
            fs.rank_repo_url(1, 10, forbidden)
        ready, why = fs.rank_repo_ready(forbidden)
        self.assertFalse(ready)
        self.assertIn("github_profile_extractor", why)

    def test_readiness_is_judged_without_writing_anything(self):
        source = inspect.getsource(fs.rank_repo_ready)
        self.assertIn("ls-remote", source)
        self.assertIn("--dry-run", source)

        remote = tempfile.mkdtemp(prefix="rank_ready_")
        subprocess.run(["git", "init", "--bare", "-q", remote], check=True)
        seed = tempfile.mkdtemp(prefix="rank_seed_")
        subprocess.run(["git", "init", "-q", seed], check=True)
        for args in (("config", "user.email", "t@example.invalid"),
                     ("config", "user.name", "T")):
            git(seed, *args)
        with open(os.path.join(seed, "f.txt"), "w", encoding="utf-8") as handle:
            handle.write("x\n")
        git(seed, "add", "f.txt")
        git(seed, "commit", "-q", "-m", "seed")
        git(seed, "push", "-q", remote, "HEAD:refs/heads/master")

        kept = fs.PROJECT_DIR
        here = os.getcwd()
        try:
            fs.PROJECT_DIR = seed
            os.chdir(seed)
            ready, why = fs.rank_repo_ready(remote)
            self.assertTrue(ready, why)
            missing = os.path.join(tempfile.mkdtemp(prefix="rank_none_"),
                                   "nope.git")
            ready, why = fs.rank_repo_ready(missing)
            self.assertFalse(ready)
            self.assertIn("cannot read", why)
        finally:
            os.chdir(here)
            fs.PROJECT_DIR = kept
        # the probe left no branch behind
        self.assertNotEqual(
            git(remote, "rev-parse", "--verify", "--quiet",
                "refs/heads/_write_probe").returncode, 0)

    def test_url_generation_is_separate_from_source_synchronisation(self):
        self.assertIs(fs.RANK_REPO_ENABLED, False)
        self.assertNotIn("rank_repo_url", inspect.getsource(fs.git_sync))
        self.assertEqual(fs.total_scripts, 10)


if __name__ == "__main__":
    unittest.main()
