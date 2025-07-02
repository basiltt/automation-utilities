# Developer : Basil T T(basil.tt@hpe.com)
# Created Date: 13th March 2024
# Updated  Date: 2nd May 2024
# Version: 1.5
# Developed with Python 3.10 and Selenium 4.18.1
#######################################

import atexit
import enum
import os
import threading
import time
from contextlib import suppress
from dataclasses import dataclass
from typing import Union, Optional, List, Dict

import loguru
# from loguru import logger
from selenium import webdriver
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
)
from selenium.webdriver import Keys, ActionChains
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.remote.webelement import (
    WebElement,
)
from selenium.webdriver.support import (
    expected_conditions as EC,
)
from selenium.webdriver.support.select import (
    Select,
)
from selenium.webdriver.support.ui import (
    WebDriverWait,
)

# Defining constants
# DEFAULT_WAIT_TIME = 180
DEFAULT_WAIT_TIME = 300
SUCCESS_MESSAGE = {
    "click": "successfully clicked on the element",
    "set_text": "successfully set the text on the element",
    "switch_to_frame": "successfully switched to the frame element",
    "get_text": "successfully fetched the text from the element",
    "set_text_enter": "successfully set the text "
                      "and pressed enter on the element",
}


# Defining custom exceptions
class WebElementNotFoundError(Exception):
    """
    Custom Exception to raise when the WebElement not found
    """

    pass


class SelectionNotFoundError(Exception):
    """
    Custom Exception to raise when the selection not found
    """

    pass


class ParameterMissingError(Exception):
    """
    Custom Exception to raise when the parameter missing
    """

    pass


class StopExecution(Exception):
    """
    Custom Exception to raise when the execution needs to be stopped
    """

    pass


# Create a data class named as Data with attributes data and status
@dataclass
class Data:
    data: str
    status: bool


class Browser(enum.Enum):
    """Enum for first-class browser support."""

    CHROME = "chrome"
    FIREFOX = "firefox"
    EDGE = "edge"


def _time_left(start: float, timeout: int | float) -> float:
    """Return seconds remaining before the absolute timeout expires."""
    remaining: float = float(timeout) - (time.time() - start)
    return max(0.0, remaining)


class WebActions:
    def __init__(
            self,
            *,
            browser: str | Browser = Browser.CHROME,
            chrome_path: str | None = None,
            chrome_driver_path: str | None = None,
            firefox_path: str | None = None,
            firefox_driver_path: str | None = None,
            edge_path: str | None = None,
            edge_driver_path: str | None = None,
            download_path: str | None = None,
            proxy_address: str | None = None,
            proxy_port: str | None = None,
            arguments: list[str] | None = None,
            experimental_options: dict | None = None,
            start_maximized: bool = True,
            driver=None,
            **kwargs,
    ) -> None:
        """
        Create a WebActions instance.

        Parameters
        ----------
        browser : str | Browser, optional
            Target browser – "chrome", "firefox", or "edge".
        *_path / *_driver_path : str | None
            Binary & driver-service paths for individual browsers.
        download_path : str | None
            Absolute or relative directory for downloads.
        proxy_address / proxy_port : str | None
            HTTP(S) proxy host & port.
        arguments : list[str] | None
            Extra command-line flags for the browser.
        experimental_options : dict | None
            Extra Selenium options (Chrome / Edge only).
        start_maximized : bool
            Whether to maximize window after launch.
        driver : Selenium WebDriver, optional
            Re-use an existing driver instead of creating one.
        **kwargs
            Unsupported keys are ignored except:
              • raise_exception : bool – re-throw element errors.
              • logger          : logging.Logger compatible.

        Raises
        ------
        FileNotFoundError
            For invalid driver/binary/download paths.
        """
        super().__init__()
        self.logger = kwargs.pop("logger", loguru.logger)
        self.raise_exception = kwargs.pop("raise_exception", False)

        # Normalise the browser enum
        self.browser: Browser = (
            browser
            if isinstance(browser, Browser)
            else Browser(browser.lower())
        )
        self.driver = None

        self.driver = driver or self._get_web_driver(
            chrome_path=chrome_path,
            chrome_driver_path=chrome_driver_path,
            firefox_path=firefox_path,
            firefox_driver_path=firefox_driver_path,
            edge_path=edge_path,
            edge_driver_path=edge_driver_path,
            download_path=download_path,
            proxy_address=proxy_address,
            proxy_port=proxy_port,
            arguments=arguments,
            experimental_options=experimental_options,
        )

        if self.driver and start_maximized:
            try:
                self.driver.maximize_window()
            except Exception:  # driver may not support it early in startup
                pass

        atexit.register(self.quit)

    def __del__(self):
        """
        A special method that gets called when the object is deleted.
        It prints a message and then calls the 'quit' method.
        """
        print("Deleting the WebActions object")
        if hasattr(self, "driver"):
            try:
                self.quit()
            except Exception:
                pass

    @staticmethod
    def _get_find_method(element: str) -> str:
        """
        Method to find the passed element is an ID or XPATH
        :param element: an XPATH or ID of the WebElement as string
        :return: element type as string
        """
        if element.startswith(r"//") or element.startswith(r"("):
            return "XPATH"
        return "ID"

    def _get_element_if_exist(
            self,
            element: str,
            max_wait_time: int = DEFAULT_WAIT_TIME,
            *,
            log_exception: bool = False,
            name: str | None = None,
            is_clickable: bool = True,
    ):
        """
        Try to locate the element exactly once within `max_wait_time`.
        Never blocks longer than the timeout – even in the worst case.
        """
        locator = (getattr(By, self._get_find_method(element)), element)
        start = time.time()
        cond = (
            EC.element_to_be_clickable(locator)
            if is_clickable
            else EC.presence_of_element_located(locator)
        )

        try:
            remaining = _time_left(start, max_wait_time)
            if remaining == 0:  # nothing left – abort early
                raise TimeoutException

            return WebDriverWait(
                self.driver, remaining, poll_frequency=0.25
            ).until(cond)
        except Exception as exc:  # noqa: bare except okay – re-raised below
            if log_exception:
                self.logger.error(
                    f'Element "{name or element}" not found in {max_wait_time}s',
                    exc_info=False,
                )
            if self.raise_exception:
                raise WebElementNotFoundError(str(exc)) from exc
            return None

    def set_zoom(self, percent: int) -> None:
        """
        Zoom the page to a given percentage (e.g. 70 for 70%).
        """
        script = f"document.body.style.zoom = '{percent}%';"
        self.driver.execute_script(script)
        self.logger.debug(f"Page zoom set to {percent}%")

    # ----------------------------------------------------------------------
    #  🔍  check_element_exist – never raises unless *caller* asks for it
    # ----------------------------------------------------------------------
    def check_element_exist(
            self,
            element,
            max_wait_time: int = DEFAULT_WAIT_TIME,
            *,
            log_exception: bool = False,
            name: str | None = None,
            enable_logging: bool = True,
            raise_exception: bool = False,  # ← caller-controlled
    ):
        """
        Return the element (or None) without throwing,
        *unless* `raise_exception=True` is passed explicitly.
        Instance-wide `self.raise_exception` is ignored.
        """
        if enable_logging:
            self.logger.debug(f'Checking element "{name or element}"')

        start = time.time()  # noqa
        original_flag = self.raise_exception  # ← save
        self.raise_exception = False  # ← suppress inside helper
        try:
            elem = self._get_element_if_exist(
                element,
                max_wait_time=max_wait_time,
                log_exception=log_exception,
                name=name,
                is_clickable=False,
            )
        finally:
            self.raise_exception = (
                original_flag  # ← restore regardless of outcome
            )

        if elem:
            if enable_logging:
                self.logger.debug(f'Element "{name or element}" found')
            return elem

        # timeout has expired – decide whether to raise
        if raise_exception:
            raise WebElementNotFoundError(
                f'Element "{name or element}" not found in {max_wait_time}s'
            )
        return None

    # def check_element_exist(
    #         self,
    #         element,
    #         max_wait_time: int = DEFAULT_WAIT_TIME,
    #         *,
    #         log_exception: bool = False,
    #         name: str | None = None,
    #         enable_logging: bool = True,
    #         raise_exception: bool = False,
    # ):
    #     start = time.time()
    #     if enable_logging:
    #         self.logger.debug(f'Checking element "{name or element}"')
    #
    #     elem = self._get_element_if_exist(
    #         element,
    #         max_wait_time=max_wait_time,
    #         log_exception=log_exception,
    #         name=name,
    #         is_clickable=False,
    #     )
    #
    #     if elem:
    #         if enable_logging:
    #             self.logger.debug(f'Element "{name or element}" found')
    #         return elem
    #
    #     # guarantee timeout did not overshoot
    #     _time_left(start, max_wait_time)  # noqa – call only for side-effect
    #
    #     if raise_exception:
    #         raise WebElementNotFoundError(
    #             f'Element "{name or element}" not found in {max_wait_time}s'
    #         )
    #     return None

    def wait_until_element_disappears_by_css_selector(
            self,
            selector: str,
            max_wait_time=DEFAULT_WAIT_TIME,
    ) -> None:
        """
        Method to wait until the element with the given CSS
        selector disappears.
        :param selector: The CSS selector of the element.
        :param max_wait_time: The maximum time to wait for the element
         to disappear.
        :return: None
        """
        WebDriverWait(self.driver, max_wait_time).until(
            EC.invisibility_of_element_located((By.CSS_SELECTOR, selector))
        )

    def wait_until_element_disappears(
            self,
            element,
            max_wait_time: int = DEFAULT_WAIT_TIME,
            *,
            name: str | None = None,
    ):
        log = name or element
        start = time.time()
        locator = (getattr(By, self._get_find_method(element)), element)

        self.logger.debug(f'Waiting for "{log}" to disappear')
        with suppress(TimeoutException):
            WebDriverWait(
                self.driver, max_wait_time, poll_frequency=0.25
            ).until(EC.invisibility_of_element_located(locator))
        # Either it disappeared or we are out of time – both are acceptable here
        if _time_left(start, max_wait_time) == 0:
            self.logger.debug(f'"{log}" did NOT disappear in {max_wait_time}s')
        else:
            self.logger.debug(f'"{log}" disappeared')

    def get_all_select_options(
            self,
            element,
            max_wait_time=DEFAULT_WAIT_TIME,
            value_type=None,
            name: Optional[str] = None,
            is_clickable: bool = True,
    ) -> List[str]:
        """
        Get all the select options from the given element with optional parameters for max wait time, value type, and name.

        Parameters:
            element: The element from which to get all select options.
            max_wait_time: The maximum time to wait for the element to be present (default is DEFAULT_WAIT_TIME).
            value_type: The type of value to retrieve from the select options.
            name: Optional name of the element.
            is_clickable: A flag indicating if the element needs to be clickable.

        Returns:
            A list of strings containing all the select options.
        """
        log_text = name if name else element
        self.logger.debug(
            f'Getting all options from the select element "{log_text}"'
        )
        self.element = self._get_element_if_exist(
            element, max_wait_time, is_clickable=is_clickable
        )
        if self.element:
            select_element = Select(self.element)
            if value_type == "visible_text":
                options = [option.text for option in select_element.options]
            elif value_type == "value":
                options = [
                    option.get_attribute("value")
                    for option in select_element.options
                ]
            elif value_type == "index":
                options = [i for i in range(len(select_element.options))]
            elif value_type == "child_text":
                options = [
                    option.get_attribute("text")
                    for option in select_element.options
                ]
            else:
                options = [option.text for option in select_element.options]

            self.logger.debug(
                f'Successfully fetched all options from the select element "{log_text}"'
            )
            return options
        self.logger.debug(
            f'Failed to fetch all options from the select element "{log_text}"'
        )
        return []

    def get_all_child_inner_text(
            self,
            element,
            max_wait_time=DEFAULT_WAIT_TIME,
            name: Optional[str] = None,
            is_clickable: bool = True,
    ) -> List[str]:
        """
        Get all the inner text of the child elements of the specified element.

        Parameters:
            element: The element to get the inner text of its children.
            max_wait_time: The maximum time to wait for the element to be present (default is DEFAULT_WAIT_TIME).
            name: An optional name for the element.
            is_clickable: A flag indicating if the element needs to be clickable.

        Returns:
            List[str]: A list of strings containing the inner text of the child elements.
        """
        log_text = name if name else element
        self.logger.debug(
            f'Getting all inner text of the child elements of "{log_text}"'
        )
        self.element = self._get_element_if_exist(
            element, max_wait_time, is_clickable=is_clickable
        )
        if self.element:
            inner_text = [
                child.get_property("text")
                for child in self.element.find_elements(By.XPATH, "./*")
            ]
            self.logger.debug(
                f'Successfully fetched all inner text of the child elements of "{log_text}"'
            )
            return inner_text
        self.logger.debug(
            f'Failed to fetch all inner text of the child elements of "{log_text}"'
        )
        return []

    def _perform_selection_action(
            self,
            element,
            action,
            max_wait_time=DEFAULT_WAIT_TIME,
            index=None,
            visible_text=None,
            value=None,
            is_clickable: bool = True,
            name: Optional[str] = None,
    ) -> None:
        """
        Method to perform selection or deselection on WebElement
        :param element: ID or XPATH of the WebElement as string
        :param action: action to perform on the element as string.
        (AvailableActions: select_element,deselect_element)
        :param max_wait_time: maximum wait time to wait for the element
        to be available on the dom as integer
        :param index: index of the element to do the selection for the
        select by index method as integer
        :param visible_text: visible_text of the element to do the selection
        for the select by visible_text
        method as string
        :param value: value of the element to do the selection for the select
        by value method as string
        :param is_clickable: A flag indicating if the element needs to be clickable.
        :param name: An optional name for the element.
        :return: None
        """
        self.element = self._get_element_if_exist(
            element=element,
            max_wait_time=max_wait_time,
            is_clickable=is_clickable,
            name=name,
        )
        if self.element:
            select_element = Select(self.element)
            selection_params = {
                "method": "",
                "value": "",
            }
            for available_option in [
                index,
                visible_text,
                value,
            ]:
                if available_option is not None:
                    method = f'{action.removesuffix("_element")}_by_'
                    method += [
                        i for i, a in locals().items() if a == available_option
                    ][0]
                    selection_params = {
                        "method": method,
                        "value": available_option,
                    }
            try:
                action_on_element = getattr(
                    select_element,
                    selection_params.get("method"),
                )(selection_params.get("value"))

                _success_message = (
                    f'successfully {action.removesuffix("_element")}ed '
                    f'"{selection_params.get("value")}" on the element "{element}" '
                    f'{selection_params.get("method")}'
                )
                self.logger.trace(action_on_element)
                self.logger.trace(_success_message)
            except Exception as er:
                exception_message = (
                    f'desired {action.removesuffix("_element")}ion on the element '
                    f'"{element}" {selection_params.get("method")} '
                    f'"{selection_params.get("value")}" not available'
                )
                if self.logger.level == 10:
                    raise SelectionNotFoundError(
                        f"{exception_message}"
                    ).with_traceback(er.__traceback__)
                else:
                    raise SelectionNotFoundError(exception_message)

    def _perform_action(
            self,
            element,
            action,
            text=None,
            max_wait_time=DEFAULT_WAIT_TIME,
    ) -> Union[str, None]:
        """
        Method to perform the action on a WebElement
        :param element: ID or XPATH of the WebElement as string
        :param action: action to perform on the element as string.
        (AvailableActions: click, set_text,
        switch_to_frame, get_text)
        :param text: text to set in the field as string (works only on the
        supported elements like input field)
        :param max_wait_time: maximum wait time to wait for the element
        to be available on the dom as integer
        :return: text as string for the get_text action or None
        """
        self.max_wait_time = max_wait_time
        # selection_params = {"method": "", "value": ""}
        try:
            self.element = self._get_element_if_exist(element, max_wait_time)
            if self.element:
                if action == "click":
                    self.element.click()
                if action == "set_text":
                    self.element.clear()
                    self.element.send_keys(text)
                if action == "set_text_enter":
                    self.element.clear()
                    self.element.send_keys(text, Keys.ENTER)
                if action == "switch_to_frame":
                    self.driver.switch_to.frame(self.element)
                if action == "get_text":
                    return self.element.text
        except Exception as er:  # noqa
            exception_message = (
                f'element "{element}" not found within '
                f'the time frame of "{self.max_wait_time}" seconds'
            )
            if self.logger.level == 10:
                self.logger.error(
                    f"{exception_message}",
                    exc_info=True,
                )

    def click(
            self,
            element,
            max_wait_time: int = DEFAULT_WAIT_TIME,
            *,
            name: str | None = None,
            enable_logging: bool = True,
    ):
        start = time.time()
        log = name or element
        if enable_logging:
            self.logger.debug(f'Clicking on "{log}"')

        # loop only if we hit an intercept *after* we already found the element
        while _time_left(start, max_wait_time):
            try:
                remaining = _time_left(start, max_wait_time)
                elem = WebDriverWait(
                    self.driver, remaining, poll_frequency=0.25
                ).until(
                    EC.element_to_be_clickable(
                        (getattr(By, self._get_find_method(element)), element)
                    )
                )
                elem.click()
                if enable_logging:
                    self.logger.debug(f'Successfully clicked "{log}"')
                return
            except (
                    ElementClickInterceptedException,
                    StaleElementReferenceException,
            ):
                # brief pause & retry until timeout burns out
                time.sleep(0.25)
                continue
            except TimeoutException:
                break  # hard stop – we’re out of time

        if enable_logging:
            self.logger.error(f'Click on "{log}" timed out ({max_wait_time}s)')
        if self.raise_exception:
            raise WebElementNotFoundError(f'Click on "{log}" timed out')

    def set_text(
            self,
            element,
            text,
            sensitive=False,
            max_wait_time=DEFAULT_WAIT_TIME,
            validate=False,
            max_validation_attempts=3,
            validation_wait_time=2,
            name: Optional[str] = None,
            clear_text=True,
            is_clickable: bool = False,
    ) -> None:
        """
        A function to set text on a specified element with optional validation.

        Parameters:
            element: The element to set text on.
            text: The text to set.
            sensitive: A flag indicating if the text is sensitive.
            max_wait_time: The maximum time to wait for the element to be interactable.
            validate: A flag indicating if validation should be performed.
            max_validation_attempts: The maximum number of attempts for validation.
            validation_wait_time: The wait time between validation attempts.
            name: An optional name for logging purposes.
            clear_text: A flag indicating if the text field should be cleared before setting text.
            is_clickable: A flag indicating if the element needs to be clickable.

        Returns:
            None
        """
        log_text_ = text if not sensitive else "********"
        log_text = name if name else log_text_
        self.logger.debug(
            f'Setting text "{log_text_}" on element "{log_text}"'
        )
        self._perform_action(
            element,
            "set_text",
            text=text,
            max_wait_time=max_wait_time,
        )
        if validate:
            self.element = self._get_element_if_exist(
                element, max_wait_time, is_clickable=is_clickable
            )
            if self.element.get_attribute("value") != text:
                attempts = 0
                while attempts < max_validation_attempts:
                    if clear_text:
                        self.element.clear()
                    self.element.send_keys(text)
                    if self.element.get_attribute("value") == text:
                        break
                    time.sleep(validation_wait_time)
                    attempts += 1
                if attempts == max_validation_attempts:
                    self.logger.error(
                        f'Failed to set text "{log_text_}" on element "{log_text}"'
                    )
                    raise Exception(
                        f'Failed to set text "{log_text_}" on element "{log_text}"'
                    )
        self.logger.debug(
            f'Successfully set text "{log_text_}" on element "{log_text}"'
        )

    def repeat_steps_until_success(
            self,
            steps: List[Dict],
            max_wait_time=DEFAULT_WAIT_TIME,
            no_of_attempts=3,
            step_wait_time=2,
            raise_exception=True,
            return_data=False,
    ) -> Union[bool, Data]:
        """
        Repeat steps until success.
        Parameters:
            steps (List[Dict]): List of steps to be executed.
            max_wait_time (int): Maximum time to wait for each step.
            no_of_attempts (int): Number of attempts to complete all steps.
            step_wait_time (int): Time to wait between steps.
            raise_exception (bool): Whether to raise an exception if steps cannot be completed.
            return_data (bool): Whether to return the data.
        Returns:
            Union[bool, Data]: True if steps are completed, False otherwise.
        """
        data = None
        all_names = [step.get("name", step.get("element")) for step in steps]
        initial_attempts = no_of_attempts
        break_flag = False
        while no_of_attempts > 0:
            for step in steps:
                max_wait_time = step.get("max_wait_time", max_wait_time)
                no_of_attempts = step.get(
                    "no_of_attempts",
                    no_of_attempts,
                )
                if step.get("action") == "set_text":
                    self.set_text(
                        step.get("element"),
                        step.get("text"),
                        max_wait_time=max_wait_time,
                        name=step.get("name", None),
                    )
                elif step.get("action") == "set_text_enter":
                    self.set_text_enter(
                        step.get("element"),
                        step.get("text"),
                        max_wait_time=max_wait_time,
                        name=step.get("name", None),
                    )
                elif step.get("action") == "click":
                    self.click(
                        step.get("element"),
                        max_wait_time=max_wait_time,
                        name=step.get("name", None),
                    )
                elif step.get("action") == "wait_until_element_exists":
                    self.check_element_exist(
                        step.get("element"),
                        max_wait_time=max_wait_time,
                        name=step.get("name", None),
                    )
                elif step.get("action") == "wait_until_element_text_changes":
                    self.wait_until_element_text_changes(
                        step.get("element"),
                        step.get("text"),
                        max_wait_time=max_wait_time,
                        name=step.get("name", None),
                    )
                elif step.get("action") == "wait_until_text_matches":
                    self.wait_until_text_matches(
                        step.get("element"),
                        step.get("text"),
                        max_wait_time=max_wait_time,
                        name=step.get("name", None),
                    )
                elif step.get("action") == "select_element":
                    self.select_element(
                        step.get("element"),
                        max_wait_time=max_wait_time,
                        **step.get("selection_params"),
                    )
                elif step.get("action") == "deselect_element":
                    self.deselect_element(
                        step.get("element"),
                        max_wait_time=max_wait_time,
                        **step.get("selection_params"),
                    )
                elif step.get("action") == "switch_to_frame":
                    self.switch_to_frame(step.get("element"))
                elif step.get("action") == "get_text":
                    self.get_text(step.get("element"))
                elif step.get("action") == "get_url":
                    self.get_current_url()
                elif step.get("action") == "validate_text":
                    is_valid_ = self.validate_text(
                        step.get("element"),
                        step.get("text"),
                        step.get("stop_texts", None),
                        max_wait_time=max_wait_time,
                        no_of_attempts=1,
                        validation_wait_time=2,
                        name=step.get("name", None),
                    )
                    is_valid = is_valid_.status
                    data = is_valid_.data
                    if is_valid:
                        break_flag = True
                time.sleep(step_wait_time)
            no_of_attempts -= 1
            if break_flag:
                break
            if any(
                    data in stop_text
                    for stop_text in steps[-1].get("stop_texts", [])
            ):
                break_flag = True
                break

        if not break_flag:
            log_message = (
                f"Failed to complete  steps until "
                f"success after {initial_attempts} attempts on "
                f"elements '{','.join(all_names)}'"
            )
            if raise_exception:
                raise Exception(log_message)
            else:
                self.logger.warning(log_message)
                if return_data:
                    return Data(
                        data=data,
                        status=False,
                    )
                return False
        if return_data:
            return Data(
                data=data,
                status=True,
            )
        return True

    def wait_until_element_text_changes(
            self,
            element,
            text,
            max_wait_time=DEFAULT_WAIT_TIME,
            name: Optional[str] = None,
            is_clickable: bool = False,
    ) -> None:
        """
        Wait until the text on the specified element changes to the given text within a maximum wait time.

        Args:
            element: The element to wait for the text change.
            text: The text to wait for on the element.
            max_wait_time: The maximum time to wait for the text change (default is DEFAULT_WAIT_TIME).
            name: Optional name for the element (default is None).
            is_clickable: Whether the element is clickable (default is False).

        Returns:
            None
        """
        log_text = name if name else element
        self.logger.debug(
            f'Waiting for the text on element "{log_text}" to change'
        )
        # First wait until the text matches
        is_matched = self.wait_until_text_matches(
            element,
            text,
            max_wait_time=max_wait_time,
            name=name,
            is_clickable=is_clickable,
        )
        if not is_matched:
            self.logger.debug(f'Text on element "{log_text}" did not match')
            return
        # Then wait for the text to change
        while max_wait_time > 0:
            self.element = self._get_element_if_exist(
                element, max_wait_time, is_clickable=is_clickable
            )
            if self.element.text != text:
                break
            time.sleep(1)
            max_wait_time -= 1
        self.logger.debug(f'Text on element "{log_text}" changed')

    def threaded_wait_until_element_text_changes(
            self,
            element,
            text,
            max_wait_time=DEFAULT_WAIT_TIME,
            name: Optional[str] = None,
    ) -> None:
        """
        A function that waits until the text on a specified element changes to the given text.

        Parameters:
            element: The element to wait for the text change.
            text: The text to wait for on the element.
            max_wait_time: The maximum time to wait for the text change (default is DEFAULT_WAIT_TIME).
            name: An optional name for the element.

        Returns:
            None
        """
        log_text = name if name else element
        self.logger.debug(
            f'Waiting for the text on element "{log_text}" to change'
        )
        thread = threading.Thread(
            target=self.wait_until_element_text_changes,
            args=(element, text, max_wait_time, name),
        )
        thread.start()
        thread.join()

    def wait_until_text_matches(
            self,
            element,
            text,
            max_wait_time=DEFAULT_WAIT_TIME,
            name: Optional[str] = None,
            is_clickable: bool = False,
    ) -> bool:
        """
        A function that waits until the text on a specified element matches a given text.

        Parameters:
            element: The element to check for text matching.
            text: The text to match on the element.
            max_wait_time: The maximum time to wait for the text to match (default is DEFAULT_WAIT_TIME).
            name: An optional name for logging purposes.
            is_clickable: Whether the element is clickable (default is False).

        Returns:
            bool: True if the text on the element matches the given text within the max wait time, False otherwise.
        """
        log_text = name if name else element
        self.logger.debug(
            f'Waiting for the text on element "{log_text}" to match'
        )
        while max_wait_time > 0:
            self.element = self._get_element_if_exist(
                element, max_wait_time, is_clickable=is_clickable
            )
            if self.element and self.element.text == text:
                self.logger.debug(f'Text on element "{log_text}" matched')
                return True
            time.sleep(1)
            max_wait_time -= 1
        self.logger.debug(f'Text on element "{log_text}" did not match')
        return False

    def set_text_enter(
            self,
            element,
            text,
            sensitive=False,
            max_wait_time=DEFAULT_WAIT_TIME,
            validate=False,
            validation_element=None,
            max_validation_attempts=3,
            validation_wait_time=2,
            name: Optional[str] = None,
    ) -> None:
        """
        A function that sets text and presses enter on a specified element.

        Args:
            element: The element on which to set the text.
            text: The text to set on the element.
            sensitive: A boolean indicating if the text is sensitive and needs to be hidden.
            max_wait_time: The maximum wait time for the action.
            validate: A flag to indicate whether validation is required.
            validation_element: The element used for validation.
            max_validation_attempts: The maximum number of validation attempts.
            validation_wait_time: The wait time between validation attempts.
            name: An optional name for the element.

        Returns:
            None
        """
        log_text_ = text if not sensitive else "********"
        log_text = name if name else log_text_
        self.logger.debug(
            f'Setting text "{text}" and pressing enter '
            f'on element "{log_text}"'
        )
        self._perform_action(
            element,
            "set_text_enter",
            text=text,
            max_wait_time=max_wait_time,
        )
        if validate:  # INFO : validate only works if validation element passed
            if not validation_element:
                self.element = self._get_element_if_exist(
                    element, max_wait_time
                )
                # print("Got element")
                self.element.clear()
                if (
                        self.element.get_attribute("value").lower().strip()
                        != text.lower().strip()
                ):
                    attempts = 0
                    while attempts < max_validation_attempts:
                        self.element.send_keys(text, Keys.ENTER)
                        if self.element.get_attribute("value") == text:
                            break
                        time.sleep(validation_wait_time)
                        attempts += 1
                    if attempts == max_validation_attempts:
                        self.logger.error(
                            f'Failed to set text "{text}" on element "{log_text}"'
                        )
                        raise Exception(
                            f'Failed to set text "{text}" on element "{log_text}"'
                        )
            else:
                self.validate_text(
                    validation_element,
                    text,
                    max_wait_time=max_wait_time,
                    no_of_attempts=max_validation_attempts,
                    validation_wait_time=validation_wait_time,
                )

        self.logger.debug(
            f'Successfully set text "{text}" and pressed enter '
            f'on element "{log_text}"'
        )

    def validate_text(
            self,
            element,
            text,
            stop_texts: Optional[List[str]] = None,
            max_wait_time=DEFAULT_WAIT_TIME,
            no_of_attempts=3,
            validation_wait_time=2,
            name: Optional[str] = None,
    ) -> Data:
        """
        Validate the text on the given element multiple times with retries.

        :param element: The element to validate the text on.
        :param text: The text to be validated on the element.
        :param max_wait_time: Maximum time to wait for the element to be available.
        :param no_of_attempts: Number of attempts to validate the text.
        :param validation_wait_time: Time to wait between validation attempts.
        :param name: Optional name of the element for logging purposes.
        :param stop_texts: A list of texts to stop the validation.
        :return: A Data object with the validation status.
        """
        log_text = name if name else element
        self.logger.debug(f'Validating text "{text}" on element "{log_text}"')
        data = None
        attempts = 0
        while attempts < no_of_attempts:
            self.element = self._get_element_if_exist(element, max_wait_time)
            self.logger.debug(f"Text on the element: {self.element.text}")
            data = self.element.text
            if stop_texts and self.element.text in stop_texts:
                self.logger.debug(
                    f'Stopping validation as the text "{self.element.text}" '
                    f"is in the stop texts list"
                )
                return Data(
                    data=data,
                    status=False,
                )
            if data == text:
                self.logger.debug(
                    f'Successfully validated text "{text}" on element "{log_text}"'
                )
                return Data(
                    data=data,
                    status=True,
                )
            time.sleep(validation_wait_time)
            attempts += 1
        self.logger.debug(
            f'Failed to validate text "{text}" on element "{log_text}"'
        )
        return Data(
            data=data,
            status=False,
        )

    def execute_script(
            self,
            script: str,
            name: Optional[str] = None,
            *args,
    ):
        """
        Execute a given script using the driver.

        :param script: The script to be executed.
        :param name: Optional name for the script.
        :param args: Additional arguments to be passed to the script.
        :return: None
        """
        log_text = name if name else script
        self.logger.debug(f'Executing script "{log_text}"')
        self.driver.execute_script(script, *args)
        self.logger.debug(f'Successfully executed script "{log_text}"')

    def click_and_set_text(
            self,
            element,
            text,
            max_wait_time=DEFAULT_WAIT_TIME,
            name: Optional[str] = None,
    ):
        """
        Clicks on the element, sets the provided text, and logs the action.

        Args:
            element: The element to click on and set text.
            text: The text to be set in the element.
            max_wait_time: Maximum time to wait for the element (default is DEFAULT_WAIT_TIME).
            name: Optional name for logging purposes (default is None).
        """
        log_text = name if name else element
        self.logger.debug(
            f'Clicking on element "{log_text}" and setting text "{text}"'
        )
        element_ = self._get_element_if_exist(element, max_wait_time)
        action = ActionChains(self.driver)
        action.click(on_element=element_)
        action.send_keys(text, Keys.RETURN)
        time.sleep(0.8)
        action.perform()
        self.logger.debug(
            f'Successfully clicked on element "{log_text}" '
            f'and set text "{text}"'
        )

    def get_drop_down_exact_value_by_value(
            self,
            element,
            value,
            max_wait_time=DEFAULT_WAIT_TIME,
            name: Optional[str] = None,
    ):
        """
        A function to get the exact text value from a dropdown element based on the provided value.

        Parameters:
            element (str): The ID of the dropdown element.
            value (str): The value to match in the dropdown options.
            max_wait_time (int): Maximum time to wait for the element to be located (default is DEFAULT_WAIT_TIME).
            name (str, optional): An optional name parameter.

        Returns:
            str: The exact text of the option matching the provided value.
        """
        exact_pd = ""
        drop_down = Select(self.driver.find_element(By.ID, element))
        for opt in drop_down.options:
            if value == opt.get_attribute("value"):
                exact_pd = opt.get_attribute("innerText")
        return exact_pd

    def action_chain(self, actions: List[Dict]):
        """
        A method that executes a chain of actions based on the provided list of dictionaries.

        Parameters:
            actions (List[Dict]): A list of dictionaries where each dictionary represents an action to be performed.

        Returns:
            None
        """
        self.logger.debug(f"Performing action chain: {actions}")
        for action in actions:
            if action.get("action") == "click":
                self.click(action.get("element"))
            if action.get("action") == "clear":
                self.clear_text(action.get("element"))
            if action.get("action") == "set_text":
                self.set_text(
                    action.get("element"),
                    action.get("text"),
                )
            if action.get("action") == "set_text_enter":
                self.set_text_enter(
                    action.get("element"),
                    action.get("text"),
                )
            if action.get("action") == "switch_to_frame":
                self.switch_to_frame(action.get("element"))
            if action.get("action") == "get_text":
                self.get_text(action.get("element"))
            if action.get("action") == "get_url":
                self.get_current_url()
        self.logger.debug(f"Successfully performed action chain: {actions}")

    def clear_text(
            self,
            element,
            max_wait_time=DEFAULT_WAIT_TIME,
            name: Optional = None,
    ) -> None:
        """
        Clear the text from an element, optionally waiting up to a specified time.

        :param element: The element from which to clear the text.
        :param max_wait_time: The maximum time to wait for the element to be present.
        :param name: An optional name for the element.
        :return: None
        """
        log_text = name if name else element
        self.logger.debug(f'Clearing text on element "{log_text}"')
        self.element = self._get_element_if_exist(element, max_wait_time)
        if self.element:
            self.element.clear()
            self.logger.debug(
                f'Successfully cleared text on element "{log_text}"'
            )

    def switch_to_frame(
            self,
            element,
            max_wait_time=DEFAULT_WAIT_TIME,
    ) -> None:
        """
        Method to switch control to a frame
        :param element: ID or XPATH of the WebElement as string
        :param max_wait_time: maximum wait time to wait for the element
        to be available on the dom as integer
        :return: None
        """
        self._perform_action(
            element,
            "switch_to_frame",
            max_wait_time=max_wait_time,
        )

    def get_text(
            self,
            element,
            max_wait_time=DEFAULT_WAIT_TIME,
            name: Optional[str] = None,
    ) -> str:
        """
        A method to get text from a given element with an optional name, using a specified maximum wait time.

        Parameters:
            element: The element from which to fetch the text.
            max_wait_time: The maximum time to wait for the text to be fetched (default is DEFAULT_WAIT_TIME).
            name: An optional name to be used in logging (default is None).

        Returns:
            The fetched text as a string.
        """
        log_text = name if name else element
        self.logger.debug(f'Fetching text from element "{log_text}"')
        fetched_text = self._perform_action(
            element,
            "get_text",
            max_wait_time=max_wait_time,
        )
        self.logger.debug(
            f'Successfully fetched text from element "{log_text}"'
        )
        return fetched_text

    def get_element(
            self,
            element,
            max_wait_time=DEFAULT_WAIT_TIME,
            is_clickable: bool = True,
    ) -> Union[WebElement, None]:
        """
        A function that retrieves an element based on specified parameters and returns it.

            Parameters:
                element: The element to retrieve.
                max_wait_time: The maximum time to wait for the element to be available (default is DEFAULT_WAIT_TIME).
                is_clickable: A boolean indicating whether the element should be clickable (default is True).

            Returns:
                Union[WebElement, None]: The retrieved element or None if not found.
        """
        self.element = self._get_element_if_exist(
            element, max_wait_time, is_clickable=is_clickable
        )
        return self.element if self.element else None

    def select_element(
            self,
            element,
            max_wait_time=DEFAULT_WAIT_TIME,
            name: Optional[str] = None,
            is_clickable: bool = True,
            **kwargs,
    ) -> None:
        """
        A method to select an element with various options like index, visible text, or value.

        Parameters:
            element: The element to select.
            max_wait_time: The maximum wait time for the selection process. Defaults to DEFAULT_WAIT_TIME.
            name: An optional name for the element.
            is_clickable: A flag indicating if the element needs to be clickable to be selected.
            **kwargs: Additional keyword arguments for more options like index, visible text, or value.

        Returns:
            None
        """
        log_text = name if name else element
        self.logger.debug(f'Selecting element "{log_text}"')
        index = kwargs.pop("index", None)
        visible_text = kwargs.pop("visible_text", None)
        value = kwargs.pop("value", None)
        if not any([index, visible_text, value]):
            self.logger.error(
                'to select element either "index" or "visible_text" or '
                '"value" should be passed',
                exc_info=True,
            )
            raise ParameterMissingError(
                'to select element either parameter "index" or "visible_text" or '
                '"value" should be passed'
            )
        self._perform_selection_action(
            element,
            "select_element",
            max_wait_time=max_wait_time,
            index=index,
            visible_text=visible_text,
            value=value,
            name=name,
            is_clickable=is_clickable,
        )
        self.logger.debug(f'Successfully selected element "{log_text}"')

    def deselect_element(
            self,
            element,
            max_wait_time=DEFAULT_WAIT_TIME,
            **kwargs,
    ) -> None:
        """
        Method to do the deselection on supported WebElement
        :param element: ID or XPATH of the WebElement as string
        :param max_wait_time: maximum wait time to wait for the
        element to be available on the dom as integer
        :param kwargs: Optional Keyword Arguments for the
        correspondent de selection method
        Supported Keyword Arguments :-
        visible_text: visible text of the selection as string
        index: index value of the selection as integer
        value: value of the selection as string
        :return: None
        """
        index = kwargs.pop("index", None)
        visible_text = kwargs.pop("visible_text", None)
        value = kwargs.pop("value", None)
        if not any([index, visible_text, value]):
            self.logger.error(
                'to deselect element either parameter "index" or '
                '"visible_text" or "value" should be passed',
                exc_info=True,
            )
            raise ParameterMissingError(
                'to deselect element either parameter "index" or '
                '"visible_text" or "value" should be passed'
            )
        self._perform_selection_action(
            element,
            "deselect_element",
            max_wait_time=max_wait_time,
            index=index,
            visible_text=visible_text,
            value=value,
        )

    def is_enabled(
            self,
            element,
            max_wait_time=DEFAULT_WAIT_TIME,
            is_clickable: bool = True,
            name: Optional[str] = None,
    ) -> bool:
        """
        Check if the element is enabled within a specified wait time.

        Args:
            element: The element to check.
            max_wait_time: The maximum time to wait for the element to be enabled.
            is_clickable: Whether the element should also be clickable to be considered enabled.
            name: Optional name to use in logging instead of the element itself.

        Returns:
            bool: True if the element is enabled, False otherwise.
        """
        log_text = name if name else element
        self.logger.debug(f'Checking if element "{log_text}" is enabled')
        self.element = self._get_element_if_exist(
            element, max_wait_time, is_clickable=is_clickable
        )
        if self.element:
            if self.element.is_enabled():
                self.logger.debug(f'Element "{log_text}" is enabled')
                return True
        self.logger.debug(f'Element "{log_text}" is disabled')
        return False

    def is_selected(
            self,
            element,
            max_wait_time=DEFAULT_WAIT_TIME,
            is_clickable: bool = True,
            name: Optional[str] = None,
    ) -> bool:
        """
        A description of the entire function, its parameters, and its return types.

            Args:
                element: The element to check if it is selected.
                max_wait_time: The maximum time to wait for the element to be found (default is DEFAULT_WAIT_TIME).
                is_clickable: A boolean indicating whether the element needs to be clickable to be considered selected (default is True).
                name: An optional name for the element.

            Returns:
                bool: True if the element is selected, False otherwise.
        """
        log_text = name if name else element
        self.logger.debug(f'Checking if element "{log_text}" is selected')
        self.element = self._get_element_if_exist(
            element, max_wait_time, is_clickable=is_clickable
        )
        if self.element:
            if self.element.is_selected():
                self.logger.debug(f'Element "{log_text}" is selected')
                return True
        self.logger.debug(f'Element "{log_text}" is not selected')
        return False

    def navigate_to(self, url: str, name: Optional[str] = None) -> None:
        """
        A function to navigate to a specified URL, with an optional name parameter.
        Parameters:
            url: str - the URL to navigate to
            name: Optional[str] - an optional name parameter, defaults to None
        Returns:
            None
        """
        log_text = name if name else url
        self.logger.debug(f'Navigating to "{log_text}"')
        self.driver.get(url)
        self.logger.debug(f'Successfully navigated to "{log_text}"')

    def wait_for_element_to_be_visible(
            self,
            element,
            max_wait_time=DEFAULT_WAIT_TIME,
            name: Optional[str] = None,
    ) -> bool:
        """
        A function that waits for an element to be visible on the webpage.

        Parameters:
            element: The element to wait for.
            max_wait_time: The maximum time to wait for the element to be visible (default is DEFAULT_WAIT_TIME).
            name: An optional name for the element.

        Returns:
            bool: True if the element is visible, False if it is not visible.
        """
        log_text = name if name else element
        self.logger.debug(f'Waiting for element "{log_text}" to be visible')
        element_ = WebDriverWait(self.driver, max_wait_time).until(
            EC.visibility_of_element_located(
                (
                    getattr(
                        By,
                        self._get_find_method(element),
                    ),
                    element,
                )
            )
        )
        if element_:
            self.logger.debug(f'Element "{log_text}" is visible')
            return True
        self.logger.debug(f'Element "{log_text}" is not visible')
        return False

    def get_inner_html(
            self,
            element,
            max_wait_time=DEFAULT_WAIT_TIME,
            name: Optional[str] = None,
    ) -> str:
        """
        A function to get the inner HTML of a given element.

        Parameters:
            element: The element to extract inner HTML from.
            max_wait_time: Maximum time to wait for the element (default is DEFAULT_WAIT_TIME).
            name: Optional name of the element (default is None).

        Returns:
            str: The inner HTML of the element, or None if element is not found.
        """
        log_text = name if name else element
        self.logger.debug(f'Fetching inner HTML of element "{log_text}"')
        if isinstance(element, WebElement):
            self.element = element
        else:
            self.element = self._get_element_if_exist(element, max_wait_time)
        if self.element:
            inner_html = self.element.get_attribute("innerHTML")
            self.logger.debug(
                f'Successfully fetched inner HTML of element "{log_text}"'
            )
            return inner_html
        self.logger.debug(
            f'Failed to fetch inner HTML of element "{log_text}"'
        )
        return None

    def get_parent_element(
            self,
            element,
            max_wait_time=DEFAULT_WAIT_TIME,
            name: Optional[str] = None,
    ) -> WebElement:
        """
        A function to retrieve the parent element of a given element with an optional maximum wait time.

        Parameters:
            element: WebElement - The element to find the parent element for.
            max_wait_time: int - The maximum time to wait for the element to be found. Default is DEFAULT_WAIT_TIME.
            name: str, optional - A name to log for the element.

        Returns:
            WebElement: The parent element of the given element if found, otherwise None.
        """
        log_text = name if name else element
        self.logger.debug(f'Fetching parent element of element "{log_text}"')
        self.element = self._get_element_if_exist(element, max_wait_time)
        if self.element:
            parent_element = self.element.find_element(By.XPATH, "..")
            self.logger.debug(
                f'Successfully fetched parent element of element "{log_text}"'
            )
            return parent_element
        self.logger.debug(
            f'Failed to fetch parent element of element "{log_text}"'
        )
        return None

    def get_child_elements(
            self,
            element: str | WebElement,
            max_wait_time: int = DEFAULT_WAIT_TIME,
            *,
            name: str | None = None,
            is_clickable: bool = False,
    ) -> list[WebElement]:
        """
        Return every *direct* child of *element* as a list of WebElements.

        Parameters
        ----------
        element        ID / XPath string **or** an already-resolved WebElement
        max_wait_time  Wait for the parent element to appear before giving up
        name           Friendly name for cleaner log messages
        is_clickable   Pass True if the parent must be clickable before we continue
        """
        log = name or element
        self.logger.debug(f'Fetching direct children of "{log}"')

        # Get the parent element (either we already have it or look it up)
        parent = element if isinstance(element, WebElement) else \
            self._get_element_if_exist(
                element,
                max_wait_time=max_wait_time,
                is_clickable=is_clickable,
                name=name,
            )
        if not parent:
            self.logger.debug(f'Parent "{log}" not found → returning []')
            return []

        # "./*" = only first-level descendants
        child_elements = parent.find_elements(By.XPATH, "./*")
        self.logger.debug(f'Found {len(child_elements)} children under "{log}"')
        return child_elements

    def get_children_count(
            self,
            element,
            max_wait_time: int = DEFAULT_WAIT_TIME,
            *,
            name: str | None = None,
            is_clickable: bool = False,
    ) -> int:
        """
        Return the number of **direct** children an element has.
        """
        el = self._get_element_if_exist(
            element,
            max_wait_time=max_wait_time,
            is_clickable=is_clickable,
            name=name,
        )
        if not el:
            return 0
        return len(el.find_elements(By.XPATH, "./*"))

    def get_all_elements(
            self,
            locator: str,
            max_wait_time: int = DEFAULT_WAIT_TIME,
            *,
            name: str | None = None,
            is_clickable: bool = False,  # usually False for bulk queries
    ) -> list[WebElement]:
        """
        Return *every* element that matches *locator* (ID or XPath).

        Examples
        --------
        rows = wa.get_all_elements("//table[@id='tbl']/tbody/tr")
        links = wa.get_all_elements("//*[@href]", name="All links")
        """
        log = name or locator
        self.logger.debug(f'Fetching all elements matching "{log}"')

        by = getattr(By, self._get_find_method(locator))

        # Wait until at least one element appears (or timeout)
        try:
            WebDriverWait(self.driver, max_wait_time).until(
                EC.presence_of_element_located((by, locator))
            )
        except TimeoutException:
            self.logger.debug(f'No matches found for "{log}"')
            return []

        # If the caller asked for clickables, filter after retrieval
        elements = self.driver.find_elements(by, locator)
        if is_clickable:
            elements = [el for el in elements if el.is_enabled() and el.is_displayed()]

        self.logger.debug(f'Found {len(elements)} matches for "{log}"')
        return elements

    def get_elements_count(
            self,
            locator: str,
            max_wait_time: int = DEFAULT_WAIT_TIME,
            *,
            name: str | None = None,
            is_clickable: bool = False,
    ) -> int:
        """
        Shortcut that returns just the *count* of elements matching *locator*.

        Internally calls `get_all_elements()` to keep the logic in one place.
        """
        return len(
            self.get_all_elements(
                locator,
                max_wait_time=max_wait_time,
                name=name,
                is_clickable=is_clickable,
            )
        )

    def count_elements(
            self,
            locator: str,
            max_wait_time: int = DEFAULT_WAIT_TIME,
            *,
            name: str | None = None,
            is_clickable: bool = False,
    ) -> int:
        """
        Return the number of elements that match *locator* (ID or XPath).
        """
        log = name or locator
        self.logger.debug(f'Counting elements for "{log}"')

        by = getattr(By, self._get_find_method(locator))
        # Wait until at least one element (or timeout)
        try:
            WebDriverWait(self.driver, max_wait_time).until(
                EC.presence_of_element_located((by, locator))
            )
        except TimeoutException:
            self.logger.debug(f'No matches found for "{log}"')
            return 0

        total = len(self.driver.find_elements(by, locator))
        self.logger.debug(f'Found {total} matches for "{log}"')
        return total

    def accept_alert(self) -> None:
        """
        Method to accept an alert dialog in the current session.
        """
        self.driver.switch_to.alert.accept()

    def dismiss_alert(self) -> None:
        """
        Dismisses the current alert by switching to it and then dismissing it.
        """
        self.driver.switch_to.alert.dismiss()

    def scroll_to_element(
            self,
            element,
            max_wait_time=DEFAULT_WAIT_TIME,
            name: Optional[str] = None,
    ) -> None:
        """
        Scrolls to the specified element on the page.

        Parameters:
            element: The element to scroll to.
            max_wait_time: int - The maximum time to wait for the element to be found. Default is DEFAULT_WAIT_TIME.
            name: str, optional - A name to log for the element.


        Returns:
            None
        """
        log_text = name if name else element
        self.logger.debug(f"Scrolling to element {log_text}")
        if not isinstance(element, WebElement):
            self.element = self._get_element_if_exist(element)
        else:
            self.element = element
        if self.element:
            self.driver.execute_script(
                "arguments[0].scrollIntoView();",
                self.element,
            )
            self.logger.debug(f"Successfully scroll to element {log_text}")
        else:
            self.logger.debug(f"Failed to scroll to element {log_text}")

    def navigate_back(self) -> None:
        """
        Navigates the driver back to the previous page.
        No parameters.
        Returns None.
        """
        self.driver.back()

    def navigate_forward(self) -> None:
        """
        Navigates the driver forward.
        """
        self.driver.forward()

    def refresh_page(self) -> None:
        """
        Refreshes the current page.
        """
        self.driver.refresh()

    def get_current_url(self) -> str:
        """
        Return the current URL of the driver.
        """
        return self.driver.current_url

    def get_page_title(self) -> str:
        """
        Get the title of the current page.

        Returns:
            str: The title of the page.
        """
        return self.driver.title

    def add_cookie(self, cookie_dict) -> None:
        """
        Add a cookie to the current session.

        :param cookie_dict: A dictionary containing the cookie data to be added.
        :return: None
        """
        self.driver.add_cookie(cookie_dict)

    def get_cookie(self, name) -> dict:
        """
        A function that retrieves a specific cookie by name.

        Parameters:
            name (str): The name of the cookie to retrieve.

        Returns:
            dict: A dictionary containing the cookie information.
        """
        return self.driver.get_cookie(name)

    def get_all_cookies(self) -> list:
        """
        Get all the cookies from the driver.
        :return: list of dictionaries representing cookies
        """
        return self.driver.get_cookies()

    def delete_cookie(self, name) -> None:
        """
        Deletes a cookie by name using the provided name parameter.

        Parameters:
            name (str): The name of the cookie to be deleted.

        Returns:
            None
        """
        self.driver.delete_cookie(name)

    def delete_all_cookies(self) -> None:
        """
        Delete all cookies.
        """
        self.driver.delete_all_cookies()

    def switch_to_window(self, window_handle) -> None:
        """
        Switches to the specified window handle.

        :param window_handle: The handle of the window to switch to.
        :return: None
        """
        self.driver.switch_to.window(window_handle)

    def get_current_window_handle(self) -> str:
        """
        Return the current window handle.

        :param self: The instance of the class.
        :return: str - The current window handle.
        """
        return self.driver.current_window_handle

    def get_window_handles(self) -> list:
        """
        Get the window handles from the driver.

        :return: list of window handles
        """
        return self.driver.window_handles

    def switch_to_default_content(self) -> None:
        """
        Switches the driver to the default content.
        """
        self.driver.switch_to.default_content()

    def maximize_window(self) -> None:
        """
        Maximizes the window using the Selenium driver.
        """
        self.driver.maximize_window()

    def minimize_window(self) -> None:
        """
        A method to minimize the window using the driver.
        No parameters are needed.
        Returns None.
        """
        self.driver.minimize_window()

    def set_window_size(self, width, height) -> None:
        """
        Sets the window size using the specified width and height.

        Parameters:
            width (int): The width for the window size.
            height (int): The height for the window size.

        Returns:
            None
        """
        self.driver.set_window_size(width, height)

    def get_window_position(self) -> dict:
        """
        Get the position of the window.

        :return: dict
        """
        return self.driver.get_window_position()

    def set_window_position(self, x, y) -> None:
        """
        Set the position of the window to the specified coordinates.

        Parameters:
            x (int): The x-coordinate to set the window position to.
            y (int): The y-coordinate to set the window position to.

        Returns:
            None
        """
        self.driver.set_window_position(x, y)

    def close(self) -> None:
        """
        Close the current window.
        """
        self.driver.close()

    def quit(self) -> None:
        """
        A method that quits the WebActions object.
        """
        print("Quitting the WebActions object")
        if self.driver:
            self.driver.quit()
            self.driver = None

    def _get_web_driver(
            self,
            *,
            chrome_path,
            chrome_driver_path,
            firefox_path,
            firefox_driver_path,
            edge_path,
            edge_driver_path,
            download_path,
            proxy_address,
            proxy_port,
            arguments,
            experimental_options,
    ):
        """Return a ready WebDriver object for the selected browser."""
        self._validate_download_path(download_path)

        if self.browser is Browser.CHROME:
            opts = self._get_chrome_options(
                binary_path=chrome_path,
                download_path=download_path,
                proxy_address=proxy_address,
                proxy_port=proxy_port,
                arguments=arguments,
                experimental_options=experimental_options,
            )
            # If user provided a driver path, wrap it in Service:
            if chrome_driver_path:
                service = ChromeService(
                    executable_path=self._validate_driver(
                        chrome_driver_path, "Chrome driver"
                    )
                )
                return webdriver.Chrome(service=service, options=opts)

            # Otherwise, let Selenium Manager handle it—no executable_path arg:
            return webdriver.Chrome(options=opts)

        if self.browser is Browser.FIREFOX:
            opts = self._get_firefox_options(
                binary_path=firefox_path,
                download_path=download_path,
                proxy_address=proxy_address,
                proxy_port=proxy_port,
                arguments=arguments,
            )
            # 1) If a custom driver path is provided, wrap it in Service:
            if firefox_driver_path:
                service = FirefoxService(
                    executable_path=self._validate_driver(
                        firefox_driver_path, "Gecko driver"
                    )
                )
                return webdriver.Firefox(service=service, options=opts)

            # 2) Otherwise, let Selenium Manager handle driver download:
            return webdriver.Firefox(options=opts)

        if self.browser is Browser.EDGE:
            opts = self._get_edge_options(
                binary_path=edge_path,
                download_path=download_path,
                proxy_address=proxy_address,
                proxy_port=proxy_port,
                arguments=arguments,
                experimental_options=experimental_options,
            )
            if edge_driver_path:
                service = EdgeService(
                    executable_path=self._validate_driver(
                        edge_driver_path, "Edge driver"
                    )
                )
                return webdriver.Edge(service=service, options=opts)

                # 2) Otherwise, let Selenium Manager manage the driver
            return webdriver.Edge(options=opts)

        raise ValueError(f"Unsupported browser {self.browser!s}")

    # ———————————————————————————————————————————————
    # 3️⃣  _get_chrome_options
    # ———————————————————————————————————————————————
    @staticmethod
    def _get_chrome_options(
            *,
            binary_path: str | None,
            download_path: str,
            proxy_address: str | None,
            proxy_port: str | None,
            arguments: list[str] | None,
            experimental_options: dict | None,
    ) -> webdriver.ChromeOptions:
        """Return fully-configured ChromeOptions object."""
        options = webdriver.ChromeOptions()
        if binary_path:
            options.binary_location = binary_path

        # Downloads & proxy
        options.add_experimental_option(
            "prefs",
            {
                "download.default_directory": download_path,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
            },
        )
        if proxy_address and proxy_port:
            options.add_argument(
                f"--proxy-server={proxy_address}:{proxy_port}"
            )

        # Generic flags
        (arguments or []).append("--ignore-certificate-errors")
        for arg in set(arguments or []):
            options.add_argument(arg)

        # Experimental
        for k, v in (experimental_options or {}).items():
            options.add_experimental_option(k, v)

        # Reduce noisy DevTools logs
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        return options

    # ———————————————————————————————————————————————
    # 4️⃣  _get_firefox_options
    # ———————————————————————————————————————————————
    @staticmethod
    def _get_firefox_options(
            binary_path: Optional[str],
            download_path: str,
            proxy_address: Optional[str],
            proxy_port: Optional[int],
            arguments: Optional[List[str]],
    ) -> FirefoxOptions:
        """
        Return configured FirefoxOptions so that:
          • all downloads go straight into download_path
          • CSVs (and others you list) are saved without a prompt
        """
        opts = webdriver.FirefoxOptions()

        if binary_path:
            opts.binary_location = binary_path

        # 1) Tell Firefox to use a custom download directory
        opts.set_preference("browser.download.folderList", 2)
        opts.set_preference("browser.download.dir", download_path)
        opts.set_preference("browser.download.useDownloadDir", True)

        # 2) Suppress the “open or save” dialog for these MIME types
        #    add any comma-separated list you need
        opts.set_preference(
            "browser.helperApps.neverAsk.saveToDisk",
            "text/csv,application/csv,application/octet-stream",
        )

        # 3) Hide the download manager window
        opts.set_preference("browser.download.manager.showWhenStarting", False)
        opts.set_preference(
            "browser.download.manager.focusWhenStarting", False
        )
        opts.set_preference("browser.download.manager.alertOnEXEOpen", False)

        # 4) Disable the built-in PDF viewer (if you’re ever auto-downloading PDFs)
        opts.set_preference("pdfjs.disabled", True)

        # 5) Proxy (if required)
        if proxy_address and proxy_port:
            opts.set_preference("network.proxy.type", 1)
            opts.set_preference("network.proxy.http", proxy_address)
            opts.set_preference("network.proxy.http_port", int(proxy_port))
            opts.set_preference("network.proxy.ssl", proxy_address)
            opts.set_preference("network.proxy.ssl_port", int(proxy_port))

        # 6) Any extra CLI args
        for arg in arguments or []:
            opts.add_argument(arg)

        return opts

    # ———————————————————————————————————————————————
    # 5️⃣  _get_edge_options
    # ———————————————————————————————————————————————
    @staticmethod
    def _get_edge_options(
            *,
            binary_path: str | None,
            download_path: str,
            proxy_address: str | None,
            proxy_port: str | None,
            arguments: list[str] | None,
            experimental_options: dict | None,
    ) -> webdriver.EdgeOptions:
        """Return EdgeOptions with download/support flags."""
        options = webdriver.EdgeOptions()
        if binary_path:
            options.binary_location = binary_path

        # Edge inherits Chrome-style prefs
        options.add_experimental_option(
            "prefs",
            {
                "download.default_directory": download_path,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
            },
        )

        if proxy_address and proxy_port:
            options.add_argument(
                f"--proxy-server={proxy_address}:{proxy_port}"
            )

        for arg in set(arguments or []):
            options.add_argument(arg)

        for k, v in (experimental_options or {}).items():
            options.add_experimental_option(k, v)

        return options

    # ———————————————————————————————————————————————
    # 6️⃣  Utility validators
    # ———————————————————————————————————————————————
    @staticmethod
    def _validate_driver(path: str | None, label: str) -> str:
        if path and os.path.isfile(path):
            return path
        raise FileNotFoundError(f"{label} not found: {path!s}")

    @staticmethod
    def _validate_download_path(download_path: str | None) -> None:
        """Ensure download directory exists or create it."""
        if download_path is None:
            return  # Let the browser fall back to default
        download_path = (
            download_path
            if os.path.isabs(download_path)
            else os.path.abspath(download_path)
        )
        os.makedirs(download_path, exist_ok=True)
