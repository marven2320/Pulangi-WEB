-- Adds per-unit gate-opening columns to the `pulangi` table.
--
-- Run this once against the production database before deploying the
-- updated powerhouse.js, e.g.:
--   mysql -u root -p pulangi_data < migrations/add_gate_opening_columns.sql
--
-- Existing rows will get opening1/opening2/opening3 = 0 (see DEFAULT below);
-- powerhouse.js starts filling in real values from the next 1-second save
-- cycle after it is restarted.

ALTER TABLE `pulangi`
    ADD COLUMN `opening1` VARCHAR(16) NOT NULL DEFAULT '0' AFTER `freq3`,
    ADD COLUMN `opening2` VARCHAR(16) NOT NULL DEFAULT '0' AFTER `opening1`,
    ADD COLUMN `opening3` VARCHAR(16) NOT NULL DEFAULT '0' AFTER `opening2`;
